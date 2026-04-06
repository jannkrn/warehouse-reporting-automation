import os
import gc
import time
import shutil
import logging
import tempfile
from datetime import date, timedelta

import pandas as pd
import pyodbc
import win32com.client as win32


# =====================
# CONFIG
# =====================

DB_DSN = os.getenv("DB_DSN", "")
DB_USER = os.getenv("DB_USER", "")
DB_PASSWORD = os.getenv("DB_PASSWORD", "")

POSITIONS_XLSB = os.getenv("POSITIONS_XLSB", "")
LOG_DIR = os.getenv("LOG_DIR", os.path.join(tempfile.gettempdir(), "mok_logs"))

SHEET_DATA = os.getenv("SHEET_DATA", "Daten")
HISTORY_SHEET = os.getenv("HISTORY_SHEET", "Historie")
OUTPUT_SHEET = os.getenv("OUTPUT_SHEET", "Ausw MOK")

MAIL_TO = os.getenv("MAIL_TO", "recipient@example.com")
MAIL_CC = os.getenv("MAIL_CC", "cc@example.com")

LOG_PATH = os.path.join(LOG_DIR, "mok_loadonly.log")


def validate_config() -> None:
    required = {
        "DB_DSN": DB_DSN,
        "DB_USER": DB_USER,
        "DB_PASSWORD": DB_PASSWORD,
        "POSITIONS_XLSB": POSITIONS_XLSB,
    }
    missing = [key for key, value in required.items() if not value]
    if missing:
        raise RuntimeError(
            f"Missing required environment variables: {', '.join(missing)}"
        )


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
    logging.info("Starting database query for date %s", report_date)

    connection = pyodbc.connect(
        f"DSN={DB_DSN};UID={DB_USER};PWD={DB_PASSWORD};"
    )
    try:
        df = pd.read_sql(sql, connection)
        logging.info("Database query complete: %d rows, %d columns", len(df), df.shape[1])
        return df
    finally:
        connection.close()


def create_excel_app():
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.Interactive = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False

    try:
        excel.AutomationSecurity = 3
    except Exception:
        pass

    return excel


def write_dataframe_to_sheet(excel, workbook_path: str, sheet_name: str, df: pd.DataFrame) -> None:
    if "KOMMDATUM" in df.columns:
        df = df.copy()
        df["KOMMDATUM"] = pd.to_datetime(
            df["KOMMDATUM"], errors="coerce"
        ).dt.strftime("%Y-%m-%d")

    df = df.where(pd.notna(df), "")

    workbook = excel.Workbooks.Open(workbook_path, UpdateLinks=0)
    try:
        worksheet = workbook.Worksheets(sheet_name)
        worksheet.Cells.Clear()

        for col_index, col_name in enumerate(df.columns, start=1):
            worksheet.Cells(1, col_index).Value = str(col_name)

        if len(df) > 0:
            values = df.astype(str).values.tolist()
            worksheet.Range(
                worksheet.Cells(2, 1),
                worksheet.Cells(1 + len(values), len(df.columns))
            ).Value = values

        workbook.Save()
        logging.info("Workbook saved after writing source data: %s", workbook_path)
    finally:
        workbook.Close(SaveChanges=True)


def insert_formulas(excel, workbook_path: str) -> None:
    workbook = excel.Workbooks.Open(workbook_path, UpdateLinks=0)
    try:
        worksheet = workbook.Worksheets(SHEET_DATA)
        last_row = worksheet.Cells(worksheet.Rows.Count, 1).End(-4162).Row

        if last_row < 2:
            logging.info("No data rows found, formula insertion skipped")
            return

        headers = [
            "Bereich",
            "Bereichsnummer",
            "BH_VKST = POS/BH",
            "KOMMLISTEN JE AUFTRAG",
            "Vol pro Menge",
            "Picks/Pos",
            "Tour",
            "BE",
            "Ausgepackt",
            "Laufmeter in Regalen",
            "Ebenen-Höhe",
        ]

        first_col = worksheet.Range("V1").Column

        for offset, name in enumerate(headers):
            worksheet.Cells(1, first_col + offset).Value = name

        worksheet.Range(f"V2:V{last_row}").FormulaLocal = '="MOK"'
        worksheet.Range(f"W2:W{last_row}").FormulaLocal = "=1"
        worksheet.Range(f"X2:X{last_row}").FormulaLocal = '=N2&"_"&Q2&"_"&W2'
        worksheet.Range(f"Y2:Y{last_row}").FormulaLocal = '=R2&"_"&O2'
        worksheet.Range(f"Z2:Z{last_row}").FormulaLocal = "=WENN(K2<S2;L2*K2/1000;K2/S2*M2/1000)"
        worksheet.Range(f"AA2:AA{last_row}").FormulaLocal = "=WENN(AC2<S2;K2/1;K2/AC2)"
        worksheet.Range(f"AB2:AB{last_row}").FormulaLocal = "=SVERWEIS(LINKS(T2;2);Berechnung!S:U;3;0)"
        worksheet.Range(f"AC2:AC{last_row}").FormulaLocal = "=RECHTS(U2;7)*1"
        worksheet.Range(f"AD2:AD{last_row}").FormulaLocal = '=WENN(S2/AC2>1;"ja";"nein")'

        worksheet.Range(f"V2:AF{last_row}").Value = worksheet.Range(f"V2:AF{last_row}").Value

        workbook.Save()
        logging.info("Formulas inserted and converted to values")
    finally:
        workbook.Close(SaveChanges=True)


def update_history_sheet(excel, workbook_path: str, history_sheet_name: str = HISTORY_SHEET) -> None:
    workbook = excel.Workbooks.Open(workbook_path, UpdateLinks=0)
    try:
        output_ws = workbook.Worksheets(OUTPUT_SHEET)
        history_ws = workbook.Worksheets(history_sheet_name)

        workbook.RefreshAll()
        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        excel.CalculateFull()

        source_range = output_ws.Range("D3:S27")
        data = source_range.Value
        rows = source_range.Rows.Count
        cols = source_range.Columns.Count

        last_row = history_ws.Cells(history_ws.Rows.Count, 1).End(-4162).Row
        if last_row < 2:
            last_row = 1

        start_row = last_row + 1
        max_rows = history_ws.Rows.Count

        if start_row + rows - 1 > max_rows:
            raise RuntimeError(
                f"History sheet overflow: start_row={start_row}, rows={rows}, max_rows={max_rows}"
            )

        target_range = history_ws.Range(
            history_ws.Cells(start_row, 1),
            history_ws.Cells(start_row + rows - 1, cols)
        )
        target_range.Value = data

        workbook.Save()
        logging.info("History sheet updated starting at row %d", start_row)
    finally:
        workbook.Close(SaveChanges=True)


def remove_history_duplicates(excel, workbook_path: str, sheet_name: str = HISTORY_SHEET) -> None:
    workbook = excel.Workbooks.Open(workbook_path, UpdateLinks=0)
    try:
        worksheet = workbook.Worksheets(sheet_name)
        last_row = worksheet.Cells(worksheet.Rows.Count, 1).End(-4162).Row

        if last_row < 2:
            logging.info("No history rows found, duplicate removal skipped")
            return

        seen = set()
        rows_to_delete = []

        for row_idx in range(last_row, 1, -1):
            key = worksheet.Cells(row_idx, 2).Value
            if key is None or str(key).strip() == "":
                continue

            key = str(key).strip()
            if key in seen:
                rows_to_delete.append(row_idx)
            else:
                seen.add(key)

        for row_idx in rows_to_delete:
            worksheet.Rows(row_idx).Delete()

        workbook.Save()
        logging.info("Removed %d duplicate rows from history", len(rows_to_delete))
    finally:
        workbook.Close(SaveChanges=True)


def create_outlook_mail_draft(to_addr: str, cc_addr: str, subject: str, body: str, attachment_path: str) -> None:
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to_addr
    mail.CC = cc_addr
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(attachment_path)
    mail.Display()


def export_report_values_and_format(excel, src_workbook_path: str, out_xlsx_path: str) -> None:
    workbook = excel.Workbooks.Open(src_workbook_path, UpdateLinks=0)
    try:
        worksheet = workbook.Worksheets(OUTPUT_SHEET)

        workbook.RefreshAll()
        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass
        excel.CalculateFull()

        new_workbook = excel.Workbooks.Add()
        try:
            new_worksheet = new_workbook.Worksheets(1)
            new_worksheet.Name = OUTPUT_SHEET

            used_range = worksheet.UsedRange
            rows = used_range.Rows.Count
            cols = used_range.Columns.Count

            if rows < 1 or cols < 1:
                if os.path.exists(out_xlsx_path):
                    os.remove(out_xlsx_path)
                new_workbook.SaveAs(out_xlsx_path, FileFormat=51)
                logging.info("Empty export saved: %s", out_xlsx_path)
                return

            destination = new_worksheet.Range(
                new_worksheet.Cells(1, 1),
                new_worksheet.Cells(rows, cols)
            )

            destination.Value = used_range.Value
            destination.Value = destination.Value

            excel.CutCopyMode = False
            used_range.Copy()
            destination.PasteSpecial(-4122)
            excel.CutCopyMode = False

            for col_idx in range(1, cols + 1):
                new_worksheet.Columns(col_idx).ColumnWidth = worksheet.Columns(col_idx).ColumnWidth

            if os.path.exists(out_xlsx_path):
                os.remove(out_xlsx_path)

            new_workbook.SaveAs(out_xlsx_path, FileFormat=51)
            logging.info("Formatted export saved: %s", out_xlsx_path)
        finally:
            new_workbook.Close(SaveChanges=False)
    finally:
        workbook.Close(SaveChanges=False)


def acquire_share_lock(lock_path: str, timeout_sec: int = 600) -> None:
    start = time.time()

    while True:
        try:
            with open(lock_path, "x", encoding="utf-8") as file:
                file.write(f"locked_at={time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            return
        except FileExistsError:
            if time.time() - start > timeout_sec:
                raise RuntimeError(f"Share lock still active after {timeout_sec}s: {lock_path}")
            time.sleep(5)


def release_share_lock(lock_path: str) -> None:
    try:
        os.remove(lock_path)
    except OSError:
        pass


def copy_share_to_local(src_unc: str, local_path: str, retries: int = 5) -> None:
    for attempt in range(retries):
        try:
            shutil.copy2(src_unc, local_path)
            return
        except Exception:
            if attempt == retries - 1:
                raise
            time.sleep(3)


def atomic_replace_on_share(local_file: str, share_target: str, retries: int = 5) -> None:
    tmp_target = share_target + ".tmp"

    for attempt in range(retries):
        try:
            if os.path.exists(tmp_target):
                os.remove(tmp_target)
            shutil.copy2(local_file, tmp_target)
            os.replace(tmp_target, share_target)
            return
        except Exception:
            if attempt == retries - 1:
                raise
            time.sleep(5)


def cleanup_stale_files(xlsb_path: str, lock_max_age_sec: int = 2 * 60 * 60) -> None:
    tmp_path = xlsb_path + ".tmp"
    try:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
    except Exception:
        pass

    lock_path = xlsb_path + ".lock"
    try:
        if os.path.exists(lock_path):
            age = time.time() - os.path.getmtime(lock_path)
            if age > lock_max_age_sec:
                os.remove(lock_path)
    except Exception:
        pass


def main() -> None:
    setup_logging()
    validate_config()

    logging.info("START: local-copy reporting workflow")
    logging.info("Workbook path: %s", POSITIONS_XLSB)
    logging.info("Script path: %s", os.path.abspath(__file__))

    cleanup_stale_files(POSITIONS_XLSB)

    lock_path = POSITIONS_XLSB + ".lock"
    local_workbook = os.path.join(tempfile.gettempdir(), "positions_report_work.xlsb")

    report_date = pick_business_date()
    safe_date = report_date.replace("-", "")
    attachment_path = os.path.join(
        tempfile.gettempdir(),
        f"positions_report_{safe_date}.xlsx"
    )

    acquire_share_lock(lock_path, timeout_sec=600)

    excel = None
    try:
        logging.info("Copying workbook from share to local temp file")
        copy_share_to_local(POSITIONS_XLSB, local_workbook)

        df = fetch_data(report_date)

        excel = create_excel_app()

        logging.info("Writing source data")
        write_dataframe_to_sheet(excel, local_workbook, SHEET_DATA, df)

        logging.info("Applying formulas")
        insert_formulas(excel, local_workbook)

        logging.info("Updating history sheet")
        update_history_sheet(excel, local_workbook, HISTORY_SHEET)

        logging.info("Removing duplicate history entries")
        remove_history_duplicates(excel, local_workbook, HISTORY_SHEET)

        logging.info("Exporting formatted report")
        export_report_values_and_format(excel, local_workbook, attachment_path)

        subject = f"Positionsauswertung Vortag MOK ({report_date}) 0-24 Uhr"
        body = "Automatically generated report draft. Please review before sending."
        create_outlook_mail_draft(MAIL_TO, MAIL_CC, subject, body, attachment_path)

        logging.info("Replacing workbook on share")
        atomic_replace_on_share(local_workbook, POSITIONS_XLSB)
        logging.info("Workbook successfully replaced on share")

    except Exception:
        logging.exception("Pipeline execution failed")
        raise

    finally:
        if excel is not None:
            try:
                excel.DisplayAlerts = False
                excel.Quit()
            except Exception:
                pass

            try:
                del excel
            except Exception:
                pass

            gc.collect()
            time.sleep(2)

        release_share_lock(lock_path)

        try:
            os.remove(local_workbook)
        except OSError:
            pass

        logging.info("Finished. Log path: %s", LOG_PATH)


if __name__ == "__main__":
    main()