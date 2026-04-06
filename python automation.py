import os
import logging
import tempfile
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import pyodbc


# =====================
# CONFIG
# =====================

DB_DSN="YOUR_DSN_NAME"
DB_USER="OUR_USERNAME"
DB_PASSWORD="YOUR_PASSWORD"

EXPORT_DIR = "\\server\\share\\folder\\exports"
POSITIONS_XLSB="\\server\share\folder\Positionsauswertung_MOK_v3.xlsb"
LOG_DIR="\\server\share\folder\logs"
HISTORY_FILE = "\\server\share\folder\history\\positions_history.parquet"
LOOKUP_FILE = "\\server\share\folder\lookups\\tour_laufmeter.xlsx"
SHEET_DATA="Daten"
HISTORY_SHEET="Historie"
OUTPUT_SHEET="Ausw MOK"

MAIL_TO="to@example.com"
MAIL_CC="cc@example.com"

LOG_PATH = os.path.join(LOG_DIR, "mok_stage3.log")


# =====================
# SETUP
# =====================



def setup_logging() -> None:
    os.makedirs(LOG_DIR, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(LOG_PATH, encoding="utf-8"),
            logging.StreamHandler(),
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

    df["KOMMDATUM"] = pd.to_datetime(df["KOMMDATUM"], errors="coerce").dt.date

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

    df["Vol_pro_Menge"] = pd.NA
    cond = menge < ove
    df.loc[cond, "Vol_pro_Menge"] = vol_stck[cond] * menge[cond] / 1000
    df.loc[~cond, "Vol_pro_Menge"] = (menge[~cond] / ove[~cond]) * vol_ove[~cond] / 1000

    df["Ebenen_Hoehe"] = pd.to_numeric(
        df["EBENE"].astype("string").str[-7:],
        errors="coerce"
    )

    df["Picks_Pos"] = pd.NA
    cond_picks = df["Ebenen_Hoehe"] < ove
    df.loc[cond_picks, "Picks_Pos"] = menge[cond_picks]
    df.loc[~cond_picks, "Picks_Pos"] = menge[~cond_picks] / df.loc[~cond_picks, "Ebenen_Hoehe"]

    df["Ausgepackt"] = ((ove / df["Ebenen_Hoehe"]) > 1).map({True: "ja", False: "nein"})

    if lookup_df is not None and not lookup_df.empty:
        temp = df.copy()
        temp["TOUR_KEY"] = temp["TOUR"].astype("string").str[:2]
        lookup_df = lookup_df.copy()
        lookup_df["TOUR_KEY"] = lookup_df["TOUR_KEY"].astype("string")
        temp = temp.merge(
            lookup_df[["TOUR_KEY", "Laufmeter_in_Regalen"]],
            on="TOUR_KEY",
            how="left"
        )
        df["Laufmeter_in_Regalen"] = temp["Laufmeter_in_Regalen"]
    else:
        df["Laufmeter_in_Regalen"] = pd.NA

    return df


def load_lookup_table(path: str) -> pd.DataFrame | None:
    if not path:
        logging.info("No lookup file configured, Laufmeter lookup skipped")
        return None

    p = Path(path)
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
# HISTORY
# =====================

def build_history_payload(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ersetzt die bisherige Excel-Historie-Logik.
    Wichtig: Das hier ist ein Python-Modell.
    Wenn deine alte Historie fachlich anders war, musst du diese Aggregation anpassen.
    """
    hist = df.copy()

    hist["HIST_ID"] = (
        hist["KOMMDATUM"].astype("string")
        + "_"
        + hist["KOMMLISTE"].astype("string")
        + "_"
        + hist["VKST"].astype("string")
        + "_"
        + hist["Bereich"].astype("string")
    )

    grouped = (
        hist.groupby(
            ["HIST_ID", "KOMMDATUM", "Bereich", "Bereichsnummer", "VKST", "KOMMLISTE"],
            dropna=False,
            as_index=False
        )
        .agg(
            Positionen=("ARTIKEL", "count"),
            Picks=("Picks_Pos", "sum"),
            Menge=("MENGE", "sum"),
            Volumen_m3=("Vol_pro_Menge", "sum"),
            Behaelter=("BEHAELTER", "nunique"),
            Auftraege=("AUFTRAG", "nunique"),
        )
    )

    return grouped


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


# =====================
# REPORT
# =====================

def build_report_tables(df: pd.DataFrame, history_df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    Ersetzt die implizite Logik des Excel-Blatts 'Ausw MOK'.
    Die Tabellen kannst du fachlich anpassen, bis sie deiner bisherigen Auswertung entsprechen.
    """
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

    history_overview = (
        history_df.groupby(["KOMMDATUM", "Bereich"], as_index=False)
        .agg(
            Positionen=("Positionen", "sum"),
            Picks=("Picks", "sum"),
            Menge=("Menge", "sum"),
            Volumen_m3=("Volumen_m3", "sum"),
            Behaelter=("Behaelter", "sum"),
            Auftraege=("Auftraege", "sum"),
        )
        .sort_values(["KOMMDATUM", "Bereich"])
    )

    return {
        "Detaildaten": detail,
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

    new_history = build_history_payload(enriched_df)
    existing_history = load_history(HISTORY_FILE)
    full_history = append_and_deduplicate_history(existing_history, new_history)
    save_history(full_history, HISTORY_FILE)

    report_tables = build_report_tables(enriched_df, full_history)
    export_report(report_tables, output_path)

    logging.info("Finished successfully")
    logging.info("Output file: %s", output_path)
    logging.info("History file: %s", HISTORY_FILE)


if __name__ == "__main__":
    main()
