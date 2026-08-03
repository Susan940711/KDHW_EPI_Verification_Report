from io import BytesIO

import pandas as pd

try:
    import openpyxl
except ModuleNotFoundError:
    openpyxl = None


def load_excel_preview(uploaded_file) -> dict:
    if openpyxl is None:
        raise RuntimeError("Missing dependency 'openpyxl'. Add it to requirements.txt and redeploy.")

    workbook = openpyxl.load_workbook(BytesIO(uploaded_file.getvalue()), data_only=True)
    sheet_names = workbook.sheetnames
    first_sheet = workbook[sheet_names[0]]
    rows = list(first_sheet.iter_rows(values_only=True))

    if not rows:
        preview = pd.DataFrame()
    else:
        preview = pd.DataFrame(rows[1:], columns=rows[0])

    return {"sheet_names": sheet_names, "preview": preview}
