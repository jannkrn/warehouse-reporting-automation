import os
import logging
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import pyodbc
import win32com.client as win32


# =====================
# CONFIG
# =====================

DB_DSN = "YOUR_DSN_NAME"
DB_USER = "OUR_USERNAME"
DB_PASSWORD = "YOUR_PASSWORD"

EXPORT_DIR = "\\server\\share\\folder\\exports"
POSITIONS_XLSB = "\\server\\share\\folder\\Positionsauswertung_MOK_v3.xlsb"
LOG_DIR = "\\server\\share\\folder\\logs"
HISTORY_FILE = "\\server\\share\\folder\\history\\positions_history.parquet"
LOOKUP_FILE = "\\server\\share\\folder\\lookups\\tour_laufmeter.xlsx"
SHEET_DATA = "Daten"
HISTORY_SHEET = "Historie"
OUTPUT_SHEET = "Ausw MOK"

MAIL_TO = "to@example.com"
MAIL_CC = "cc@example.com"

LOG_PATH = os.path.join(LOG_DIR, "mok_stage3.log")


# =====================
# SETUP
# =====================

def setup_logging() -> None:
    os.makedirs(LOG_DIR, exist_ok=True)  # Stelle sicher, dass das Log-Verzeichnis existiert, bevor wir versuchen, darin zu schreiben
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(LOG_PATH, encoding="utf-8"),  # in Datei
            logging.StreamHandler(),  # in Konsole
        ],
    )


# =====================
# DATE / SQL
# =====================

def pick_business_date() -> str:
    today = date.today()
    delta = 3 if today.weekday() == 0 else 1
    return (today - timedelta(days=delta)).strftime("%Y-%m-%d")


def build_sql(report_date: str) -> str:
    return f"""
SELECT DISTINCT
    DATE(TO_DATE(RPAD(T1.TPDTK,8,'0'),'YYYYMMDD')) AS KOMMDATUM,
    CAST(SUBSTR(LPAD(T1.TPTIK,6,'0'),1,2) AS INTEGER) AS KOMMSTUNDE,
    SUBSTR(LPAD(T1.TPTIK,6,'0'),3,2) AS KOMMMINUTE,
    T1.TPVORT AS ORT,
    T1.TPVBER AS BEREICH,
    T1.TPVREG AS REGAL,
    T1.TPVHOR AS FACH,
    T1.TPVVER AS EBENE,
    T1.TPIDEN AS ARTIKEL,
    T2.TZBEZ1 AS BEZEICHNUNG,
    T1.TPBMEN AS MENGE,
    T2.TZCM3 AS VOLUMENSTCK,
    T2.TZ1CM3 AS VOLUMENOVE,
    T1.TPBENR AS BEHAELTER,
    T1.TPNRKS AS KOMMLISTE,
    T1.TPPKNR AS PACKSTUECK,
    T1.TPKDNR AS VKST,
    T1.TPANR1 AS AUFTRAG,
    T2.TZVPE AS OVE,
    T1.TPTOUR AS TOUR,
    T2.TZBEZ2 AS BE
FROM R2MDATV8.PHISTTP T1
LEFT JOIN R2MDATV8.PBFSTSS T2
    ON T1.TPIDEN = T2.TZIDEN
    AND T1.TPFIRM = T2.TZFIRM
    AND T1.TPKONZ = T2.TZKONZ
WHERE T1.TPFIRM = '002'
  AND T1.TPKONZ = '999'
  AND T1.TPVORT = 'GSL'
  AND T1.TPVBER NOT LIKE 'V%'
  AND T1.TPRCDE = 'OK'
  AND T1.TPDTK > '0'
  AND T1.TPDTK = VARCHAR_FORMAT(DATE('{report_date}'),'YYYYMMDD')
GROUP BY
    DATE(TO_DATE(RPAD(T1.TPDTK,8,'0'),'YYYYMMDD')),
    CAST(SUBSTR(LPAD(T1.TPTIK,6,'0'),1,2) AS INTEGER),
    SUBSTR(LPAD(T1.TPTIK,6,'0'),3,2),
    T1.TPVORT,
    T1.TPVBER,
    T1.TPVREG,
    T1.TPVHOR,
    T1.TPVVER,
    T1.TPIDEN,
    T2.TZBEZ1,
    T1.TPBMEN,
    T2.TZCM3,
    T2.TZ1CM3,
    T1.TPBENR,
    T1.TPNRKS,
    T1.TPPKNR,
    T1.TPKDNR,
    T1.TPANR1,
    T2.TZVPE,
    T1.TPTOUR,
    T2.TZBEZ2
"""


def fetch_data(report_date: str) -> pd.DataFrame:
    sql = build_sql(report_date)
    logging.info("Starting database query for %s", report_date)
    connection = pyodbc.connect(f"DSN={DB_DSN};UID={DB_USER};PWD={DB_PASSWORD};")
    try:
        df = pd.read_sql(sql, connection)
        logging.info("Query finished: %d rows, %d columns", len(df), df.shape[1])
        return df
    finally:
        connection.close()


# =====================
# TRANSFORM
# =====================

def normalize_types(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    df["KOMMDATUM"] = pd.to_datetime(df["KOMMDATUM"], errors="coerce").dt.date  # nur Datum, keine Uhrzeit

    numeric_cols = [
        "KOMMSTUNDE", "MENGE", "VOLUMENSTCK", "VOLUMENOVE", "OVE"
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    text_cols = [
        "KOMMMINUTE", "ORT", "BEREICH", "REGAL", "FACH", "EBENE",
        "ARTIKEL", "BEZEICHNUNG", "BEHAELTER", "KOMMLISTE", "PACKSTUECK",
        "VKST", "AUFTRAG", "TOUR", "BE"
    ]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype("string").fillna("")

    return df


def add_business_columns(df: pd.DataFrame, lookup_df: pd.DataFrame | None = None) -> pd.DataFrame:
    df = df.copy()

    df["Bereich"] = "MOK"
    df["Bereichsnummer"] = 1

    df["BH_VKST_POS_BH"] = (
        df["BEHAELTER"].astype("string").fillna("")
        + "_"
        + df["AUFTRAG"].astype("string").fillna("")
        + "_"
        + df["Bereichsnummer"].astype("Int64").astype("string")
    )

    df["KOMMLISTEN_JE_AUFTRAG"] = (
        df["VKST"].astype("string").fillna("")
        + "_"
        + df["KOMMLISTE"].astype("string").fillna("")
    )

    menge = pd.to_numeric(df["MENGE"], errors="coerce")
    ove = pd.to_numeric(df["OVE"], errors="coerce")
    vol_stck = pd.to_numeric(df["VOLUMENSTCK"], errors="coerce")
    vol_ove = pd.to_numeric(df["VOLUMENOVE"], errors="coerce")

    # =WENN(K2<S2;L2*K2/1000;K2/S2*M2/1000)
    df["Vol_pro_Menge"] = pd.NA  # noch nicht berechnet
    cond = menge < ove
    df.loc[cond, "Vol_pro_Menge"] = vol_stck[cond] * menge[cond] / 1000  # cond = true -> Formel 1
    df.loc[~cond, "Vol_pro_Menge"] = (menge[~cond] / ove[~cond]) * vol_ove[~cond] / 1000  # cond = false -> Formel 2

    df["Ebenen_Hoehe"] = pd.to_numeric(
        df["EBENE"].astype("string").str[-7:],  # nimmt von jedem string letzten 7 zeichen
        errors="coerce"
    )

    # =WENN(AC2<S2;K2/1;K2/AC2)
    df["Picks_Pos"] = pd.NA
    cond_picks = df["Ebenen_Hoehe"] < ove
    df.loc[cond_picks, "Picks_Pos"] = menge[cond_picks]  # einfach Menge als Picks, wenn Ebenen_Hoehe < OVE
    df.loc[~cond_picks, "Picks_Pos"] = menge[~cond_picks] / df.loc[~cond_picks, "Ebenen_Hoehe"]  # ansonsten Menge geteilt durch Ebenen_Hoehe

    df["Ausgepackt"] = ((ove / df["Ebenen_Hoehe"]) > 1).map({True: "ja", False: "nein"})

    if lookup_df is not None and not lookup_df.empty:
        temp = df.copy()
        temp["TOUR_KEY"] = temp["TOUR"].astype("string").str[:2]
        lookup_df = lookup_df.copy()
        lookup_df["TOUR_KEY"] = lookup_df["TOUR_KEY"].astype("string")

        merge_cols = ["TOUR_KEY", "Laufmeter_in_Regalen"]
        if "Tour_Kategorie" in lookup_df.columns:
            merge_cols.append("Tour_Kategorie")

        temp = temp.merge(
            lookup_df[merge_cols],
            on="TOUR_KEY",
            how="left"
        )
        df["Laufmeter_in_Regalen"] = temp["Laufmeter_in_Regalen"]

        if "Tour_Kategorie" in temp.columns:
            df["Tour_Kategorie"] = temp["Tour_Kategorie"]
        else:
            df["Tour_Kategorie"] = pd.NA
    else:
        df["Laufmeter_in_Regalen"] = pd.NA
        df["Tour_Kategorie"] = pd.NA

    return df


def load_lookup_table(path: str) -> pd.DataFrame | None:
    if not path:
        logging.info("No lookup file configured, Laufmeter lookup skipped")
        return None

    p = Path(path)  # Path-Objekt für einfachere Handhabung von Pfaden
    if not p.exists():
        logging.warning("Lookup file not found: %s", path)
        return None

    if p.suffix.lower() == ".csv":
        df = pd.read_csv(p)
    elif p.suffix.lower() in {".xlsx", ".xls"}:
        df = pd.read_excel(p)
    else:
        raise RuntimeError(f"Unsupported lookup file format: {p.suffix}")

    expected = {"TOUR_KEY", "Laufmeter_in_Regalen"}
    missing = expected - set(df.columns)
    if missing:
        raise RuntimeError(f"Lookup file missing columns: {', '.join(sorted(missing))}")

    return df


# =====================
# REPORT
# =====================

def build_ausw_mok(df: pd.DataFrame, report_date=None) -> dict[str, pd.DataFrame]:
    """
    Baut die eigentliche Auswertung 'Ausw MOK' nach.

    Rückgabe:
    {
        "summary": Kennzahlen links,
        "stunden": Stundenmatrix 0-1 bis 16-17
    }

    WICHTIG:
    Für die Tour-Spalten wird eine Kategorie-Spalte erwartet:
    - Tour_Kategorie mit Werten wie "GK", "P", "N"
    """

    df = df.copy()

    # Nur MOK
    if "Bereich" in df.columns:
        df = df[df["Bereich"] == "MOK"].copy()

    #auf Datum filtern
    if report_date is not None:
        report_date = pd.to_datetime(report_date).date()
        df = df[df["KOMMDATUM"] == report_date].copy()

    # Hilfsspalte für Tour-Kategorie bestimmen
    if "Tour_Kategorie" in df.columns:
        tour_cat_col = "Tour_Kategorie"
    elif "AB" in df.columns:
        tour_cat_col = "AB"
    else:
        tour_cat_col = None

    # ---------------------
    # Helper
    # ---------------------
    def safe_mean(series: pd.Series): # macht aus allem, was keine Zahl ist, NaN, damit es den mean nicht kaputt macht. Wenn danach keine Werte mehr übrig sind, wird pd.NA zurückgegeben
        series = pd.to_numeric(series, errors="coerce").dropna()
        if series.empty:
            return pd.NA
        return series.mean()

    def safe_ratio(numerator: float, denominator: float): # wenn denominator 0, None oder NaN ist, wird pd.NA zurückgegeben, ansonsten das Ergebnis der Division
        if denominator in (0, None) or pd.isna(denominator):
            return pd.NA
        return numerator / denominator

    def avg_positions_per_group(data: pd.DataFrame, group_col: str): # berechnet durchschnittliche Anzahl Positionen pro Gruppe (z.B. pro BH oder pro Haltepunkt). Wenn die Gruppe nicht existiert oder keine Positionen hat, wird pd.NA zurückgegeben
        if data.empty or group_col not in data.columns:
            return pd.NA
        counts = data.groupby(group_col).size()
        if counts.empty:
            return pd.NA
        return counts.mean()

    def avg_volume_per_group(data: pd.DataFrame, group_col: str, value_col: str): # berechnet durchschnittliches Volumen pro Gruppe (z.B. durchschnittliches Volumen pro BH). Wenn die Gruppe nicht existiert oder keine Werte hat, wird pd.NA zurückgegeben
        if data.empty or group_col not in data.columns or value_col not in data.columns:
            return pd.NA
        sums = data.groupby(group_col)[value_col].sum(min_count=1)
        sums = pd.to_numeric(sums, errors="coerce").dropna()
        if sums.empty:
            return pd.NA
        return sums.mean()

    def share_true(data: pd.DataFrame, col: str, true_value: str): # berechnet den Anteil der Zeilen, bei denen in der Spalte 'col' der Wert 'true_value' steht. Wenn die Spalte nicht existiert oder keine Zeilen hat, wird pd.NA zurückgegeben
        if data.empty or col not in data.columns:
            return pd.NA
        base = data[col].astype("string")
        if len(base) == 0:
            return pd.NA
        return (base == true_value).mean()

    def share_category(data: pd.DataFrame, col: str | None, category: str): # berechnet den Anteil der Zeilen, bei denen in der Spalte 'col' der Wert 'category' steht. Wenn die Spalte nicht existiert oder keine Zeilen hat, wird pd.NA zurückgegeben
        if data.empty or col is None or col not in data.columns:
            return pd.NA
        base = data[col].astype("string")
        if len(base) == 0:
            return pd.NA
        return (base == category).mean()

    # ---------------------
    # Gesamtkennzahlen links
    # ---------------------
    total_positions = len(df)
    total_vol_pos = safe_mean(df["Vol_pro_Menge"]) if "Vol_pro_Menge" in df.columns else pd.NA
    total_picks_pos = safe_mean(df["Picks_Pos"]) if "Picks_Pos" in df.columns else pd.NA

    total_pos_bh = avg_positions_per_group(df, "BEHAELTER")
    total_pos_haltepunkt = avg_positions_per_group(df, "VKST")
    total_vol_bh = avg_volume_per_group(df, "BEHAELTER", "Vol_pro_Menge")
    total_volumenauslastung = safe_ratio(total_vol_bh, 62.5)
    total_anteil_be_ungleich_ove = share_true(df, "Ausgepackt", "ja")

    summary_df = pd.DataFrame(
        {
            "Kennzahl": [
                "MOK Positionen",
                "MOK Vol/POS",
                "MOK Picks/POS",
                "MOK POS/BH",
                "MOK POS/Haltepunkt",
                "MOK Vol/BH",
                "MOK Volumenauslastung",
                "MOK Anteil BE ungleich OVE",
            ],
            "Wert": [
                total_positions,
                total_vol_pos,
                total_picks_pos,
                total_pos_bh,
                total_pos_haltepunkt,
                total_vol_bh,
                total_volumenauslastung,
                total_anteil_be_ungleich_ove,
            ],
        }
    )

    # Stundenmatrix
    stunden_rows = []

    for von in range(0, 23):
        bis = von + 1
        hour_df = df[df["KOMMSTUNDE"] == von].copy()

        vol_pos = safe_mean(hour_df["Vol_pro_Menge"]) if "Vol_pro_Menge" in hour_df.columns else pd.NA
        picks_pos = safe_mean(hour_df["Picks_Pos"]) if "Picks_Pos" in hour_df.columns else pd.NA

        # entspricht der alten Idee aus "Berechnung"
        pos_pro_kommliste = avg_positions_per_group(hour_df, "KOMMLISTE")
        pos_haltepunkt = avg_positions_per_group(hour_df, "VKST")
        volumenauslastung = safe_ratio(
            avg_volume_per_group(hour_df, "BEHAELTER", "Vol_pro_Menge"),
            62.5
        )

        anteil_be_ungleich_ove = share_true(hour_df, "Ausgepackt", "ja")
        grosskunden = share_category(hour_df, tour_cat_col, "GK")
        pendel = share_category(hour_df, tour_cat_col, "P")
        nachzuegler = share_category(hour_df, tour_cat_col, "N")

        stunden_rows.append(
            {
                "von": von,
                "bis": bis,
                "Vol/POS": vol_pos,
                "Picks/POS": picks_pos,
                "POS pro Komm.liste": pos_pro_kommliste,
                "POS/Haltepunkt": pos_haltepunkt,
                "Volumenauslastung": volumenauslastung,
                "Anteil BE ungleich OVE": anteil_be_ungleich_ove,
                "Großkunden": grosskunden,
                "Pendel": pendel,
                "Nachzügler": nachzuegler,
            }
        )

    stunden_df = pd.DataFrame(stunden_rows)

    stunden_df["KOMMDATUM"] = report_date if report_date is not None else pd.NA
    stunden_df["Bereich"] = "MOK"
    stunden_df["HIST_ID"] = (
        stunden_df["KOMMDATUM"].astype("string")
        + "_"
        + stunden_df["Bereich"].astype("string")
        + "_"
        + stunden_df["von"].astype("string")
        + "_"
        + stunden_df["bis"].astype("string")
    )

    #runden wie in Excel-Ansicht
    value_cols_2 = [
        "Vol/POS",
        "Picks/POS",
        "POS pro Komm.liste",
        "POS/Haltepunkt",
    ]
    percent_cols = [
        "Volumenauslastung",
        "Anteil BE ungleich OVE",
        "Großkunden",
        "Pendel",
        "Nachzügler",
    ]

    for col in value_cols_2:
        if col in stunden_df.columns:
            stunden_df[col] = pd.to_numeric(stunden_df[col], errors="coerce").round(2)

    for col in percent_cols:
        if col in stunden_df.columns:
            stunden_df[col] = pd.to_numeric(stunden_df[col], errors="coerce").round(4)

    summary_df["Wert"] = pd.to_numeric(summary_df["Wert"], errors="coerce")

    return {
        "summary": summary_df,
        "stunden": stunden_df,
    }


# =====================
# HISTORY
# =====================

def build_history_payload(ausw_mok_df: pd.DataFrame) -> pd.DataFrame:
    return ausw_mok_df.copy()


def load_history(path: str) -> pd.DataFrame:
    p = Path(path)
    if not p.exists():
        return pd.DataFrame()

    if p.suffix.lower() == ".parquet":
        return pd.read_parquet(p)
    if p.suffix.lower() == ".csv":
        return pd.read_csv(p)
    if p.suffix.lower() in {".xlsx", ".xls"}:
        return pd.read_excel(p)

    raise RuntimeError(f"Unsupported history format: {p.suffix}")


def save_history(df: pd.DataFrame, path: str) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)

    if p.suffix.lower() == ".parquet":
        df.to_parquet(p, index=False)
    elif p.suffix.lower() == ".csv":
        df.to_csv(p, index=False, sep=";")
    elif p.suffix.lower() in {".xlsx", ".xls"}:
        df.to_excel(p, index=False)
    else:
        raise RuntimeError(f"Unsupported history format: {p.suffix}")


def append_and_deduplicate_history(existing: pd.DataFrame, new_data: pd.DataFrame) -> pd.DataFrame:
    if existing.empty:
        return new_data.copy()

    combined = pd.concat([existing, new_data], ignore_index=True)
    combined = combined.drop_duplicates(subset=["HIST_ID"], keep="last")
    return combined


def build_history_overview(history_df: pd.DataFrame) -> pd.DataFrame:
    if history_df.empty:
        return pd.DataFrame()

    history_overview = (
        history_df.groupby(["KOMMDATUM", "Bereich"], as_index=False)
        .agg(
            Vol_POS=("Vol/POS", "mean"),
            Picks_POS=("Picks/POS", "mean"),
            POS_pro_Kommliste=("POS pro Komm.liste", "mean"),
            POS_Haltepunkt=("POS/Haltepunkt", "mean"),
            Volumenauslastung=("Volumenauslastung", "mean"),
            Anteil_BE_ungleich_OVE=("Anteil BE ungleich OVE", "mean"),
            Grosskunden=("Großkunden", "mean"),
            Pendel=("Pendel", "mean"),
            Nachzuegler=("Nachzügler", "mean"),
        )
        .sort_values(["KOMMDATUM", "Bereich"])
    )

    return history_overview


def build_report_tables(
    df: pd.DataFrame,
    summary_df: pd.DataFrame,
    stunden_df: pd.DataFrame,
    history_df: pd.DataFrame
) -> dict[str, pd.DataFrame]:
    detail = df.copy()

    summary_by_hour = (
        detail.groupby(["KOMMDATUM", "KOMMSTUNDE"], as_index=False)
        .agg(
            Positionen=("ARTIKEL", "count"),
            Menge=("MENGE", "sum"),
            Volumen_m3=("Vol_pro_Menge", "sum"),
            Behaelter=("BEHAELTER", "nunique"),
            Auftraege=("AUFTRAG", "nunique"),
        )
        .sort_values(["KOMMDATUM", "KOMMSTUNDE"])
    )

    summary_by_area = (
        detail.groupby(["KOMMDATUM", "BEREICH"], as_index=False)
        .agg(
            Positionen=("ARTIKEL", "count"),
            Menge=("MENGE", "sum"),
            Picks=("Picks_Pos", "sum"),
            Volumen_m3=("Vol_pro_Menge", "sum"),
        )
        .sort_values(["KOMMDATUM", "BEREICH"])
    )

    history_overview = build_history_overview(history_df)

    return {
        "Detaildaten": detail,
        "Ausw_MOK_Summary": summary_df,  # Das ist die neue zentrale Auswertung links
        "Ausw_MOK_Stunden": stunden_df,  # Das ist die neue zentrale Auswertung rechts
        "Stundenreport": summary_by_hour,
        "Bereichsreport": summary_by_area,
        "Historie": history_df,
        "Historie_Uebersicht": history_overview,
    }


def export_report(report_tables: dict[str, pd.DataFrame], output_path: str) -> None:
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, sheet_df in report_tables.items():
            safe_name = sheet_name[:31]
            export_df = sheet_df.copy()
            export_df = export_df.where(pd.notna(export_df), "")
            export_df.to_excel(writer, sheet_name=safe_name, index=False)

    logging.info("Report exported: %s", output_path)


def create_outlook_mail_send(to_addr: str, cc_addr: str, subject: str, body: str, attachment_path: str) -> None:
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to_addr
    mail.CC = cc_addr
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(attachment_path)
    mail.Send()

# =====================
# MAIN
# =====================

def main() -> None:
    setup_logging()
    report_date = pick_business_date()
    safe_date = report_date.replace("-", "")
    output_path = os.path.join(EXPORT_DIR, f"positions_report_{safe_date}.xlsx")

    logging.info("Starting stage-3 pipeline for %s", report_date)

    raw_df = fetch_data(report_date)
    raw_df = normalize_types(raw_df)

    lookup_df = load_lookup_table(LOOKUP_FILE)
    enriched_df = add_business_columns(raw_df, lookup_df)

    ausw_mok = build_ausw_mok(enriched_df, report_date)

    summary_df = ausw_mok["summary"]
    stunden_df = ausw_mok["stunden"]

    new_history = build_history_payload(stunden_df)
    existing_history = load_history(HISTORY_FILE)
    full_history = append_and_deduplicate_history(existing_history, new_history)
    save_history(full_history, HISTORY_FILE)

    report_tables = build_report_tables(enriched_df, summary_df, stunden_df, full_history)
    export_report(report_tables, output_path)

    subject = f"Positionsauswertung Vortag MOK ({report_date}) 0-24 Uhr"
    body = "Automatisch generiert – bitte prüfen."
    create_outlook_mail_send(MAIL_TO, MAIL_CC, subject, body, output_path)
    logging.info("Finished successfully")
    logging.info("Output file: %s", output_path)
    logging.info("History file: %s", HISTORY_FILE)


if __name__ == "__main__":
    main()
