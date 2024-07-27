from typing import List
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
import JustEtfScrape as etf_scrape

EXCEL_PATH = "E:\\Documents\\Trade\\Trade.xlsx"  # TODO change to receive as input
SHEET_NAME = (
    "justETF Data"  # TODO change to receive as input, also change to investigate
)
TICKER_TITLE = "Isin"  # TODO change to receive as input
COLUMN_WIDTH_PADDING = 3
HEADER_COLOR = PatternFill(start_color="C2FFC2", end_color="C2FFC2", fill_type="solid")
BREAKLINE = "\n"

COLUMNS = [
    "Isin",
    "Name",
    "Tickers",
    "Currency",
    "Volatility 1 year",
    "Volatility 3 years",
    "Volatility 5 years",
    "Returns 1 year",
    "Returns 3 years",
    "Returns 5 years",
    "TER",
    "Distribution",
    "Replication",
    "Countries",
    "Sectors",
    "Fund size",
    "Number of holdings",
]


def FlattenListToString(list: List, delimiter: str) -> str:
    return delimiter.join(list)


def LoadEtfISINFromExcel(file_path: str, column_name) -> List[str]:
    column_values = pd.read_excel(
        file_path, sheet_name=SHEET_NAME, usecols=[column_name], header=0
    )[column_name].tolist()
    return column_values


def GetEtfISINInformation(etfs_isins: List[str]) -> pd.DataFrame:
    etfs_data = etf_scrape.SeleniumScrape(etfs_isins)

    etf_dataframe = pd.DataFrame(columns=COLUMNS)
    print(etf_dataframe)

    for isin in etf_isins:
        if not pd.isna(isin):
            etf_data = etfs_data[isin]
            row_data = [
                isin,
                etf_data[etf_scrape.DICT_NAME],
                FlattenListToString(etf_data[etf_scrape.DICT_TICKERS], ", "),
                etf_data[etf_scrape.DICT_CURRENCY],
                etf_data[etf_scrape.DICT_VOLATILITY][0],
                etf_data[etf_scrape.DICT_VOLATILITY][1],
                etf_data[etf_scrape.DICT_VOLATILITY][2],
                etf_data[etf_scrape.DICT_RETURNS][0],
                etf_data[etf_scrape.DICT_RETURNS][1],
                etf_data[etf_scrape.DICT_RETURNS][2],
                etf_data[etf_scrape.DICT_TER],
                etf_data[etf_scrape.DICT_DISTRIBUTION],
                etf_data[etf_scrape.DICT_REPLICATION],
                FlattenListToString(etf_data[etf_scrape.DICT_COUNTRIES], BREAKLINE),
                FlattenListToString(etf_data[etf_scrape.DICT_SECTORS], BREAKLINE),
                etf_data[etf_scrape.DICT_FUND_SIZE],
                etf_data[etf_scrape.DICT_NUM_HOLDINGS],
            ]
            new_row = pd.DataFrame([row_data], columns=COLUMNS)
            print(new_row)
            etf_dataframe = pd.concat([etf_dataframe, new_row], ignore_index=True)

            print(etf_dataframe)
    return etf_dataframe


def WriteToExcel(etfs_info: pd.DataFrame):
    with pd.ExcelWriter(
        path=EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as excel_writer:
        etfs_info.to_excel(
            excel_writer, sheet_name=SHEET_NAME, index=False, header=True
        )
        worksheet = excel_writer.sheets[SHEET_NAME]
        # Auto-fit columns and center align
        for i, column in enumerate(etfs_info.columns):
            column_len = max(etfs_info[column].astype(str).map(len).max(), len(column))
            worksheet.column_dimensions[
                worksheet.cell(row=1, column=i + 1).column_letter
            ].width = str(
                (int(column_len) + COLUMN_WIDTH_PADDING)
                / (2 if (column == COLUMNS[13] or column == COLUMNS[14]) else 1)
            )  # Reduce padding if it is Countries of Sectors columns

            for cell in worksheet[worksheet.cell(row=1, column=i + 1).column_letter]:
                cell.alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )

        # Make the header row thicker
        worksheet.row_dimensions[
            1
        ].height = 30  # Adjust as needed for the desired height
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
            cell.fill = HEADER_COLOR


if __name__ == "__main__":
    etf_isins = LoadEtfISINFromExcel(EXCEL_PATH, TICKER_TITLE)

    etfs_info = GetEtfISINInformation(etf_isins)
    WriteToExcel(etfs_info)
