import re
from typing import List
import pandas as pd

import justetf_scraping.overview as justetf

EXCEL_PATH = "E:\Documents\Trade\Test.xlsx"  # TODO change to receive as input
SHEET_NAME = "ETF Overview"  # TODO change to receive as input


def LoadEtfTickersFromExcel(file_path: str, column_name) -> List[str]:
    column_values = pd.read_excel(
        file_path, sheet_name=SHEET_NAME, usecols=[column_name], header=0
    )[column_name].tolist()
    return column_values


def GetTickerInformation(etf_df: pd.DataFrame, etf_tickers: List[str]):
    cleaned_tickers = [re.sub(r"\..*", "", ticker) for ticker in etf_tickers]
    filtered_etfs = etfs.loc[etf_df["ticker"].isin(cleaned_tickers)]
    print(filtered_etfs)


if __name__ == "__main__":
    etfs = justetf.load_overview(enrich=True)
    etf_tickers = LoadEtfTickersFromExcel(EXCEL_PATH, "ETF TICKER")
    print(etf_tickers)
    GetTickerInformation(etfs, etf_tickers)

    # with pd.ExcelWriter(path=EXCEL_PATH) as excel_writer:
    #     df.to_excel(sheet_name='Test', excel_writer=excel_writer)
