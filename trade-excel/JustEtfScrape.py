from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from typing import List, Dict
import pandas as pd

CHROMEDRIVER_PATH = "C:\\Users\\Hugo\\Documents\\chromedriver-win64\\chromedriver.exe"
BRAVE_PATH = "C:/Program Files/BraveSoftware/Brave-Browser/Application/brave.exe"
URL_BASE = "https://www.justetf.com/en/etf-profile.html?isin="
ID_TO_CHECK = "etf-title"

DICT_NAME = "name"
DICT_TICKERS = "tickers"
DICT_CURRENCY = "currency"
DICT_VOLATILITY = "volatility"
DICT_RETURNS = "returns"
DICT_TER = "ter"
DICT_DISTRIBUTION = "distribution"
DICT_REPLICATION = "replication"
DICT_COUNTRIES = "countries"
DICT_SECTORS = "sectors"
DICT_FUND_SIZE = "fund_size"
DICT_NUM_HOLDINGS = "num_holdings"


def SoupPrettyPrint(soup: BeautifulSoup):
    print(soup.prettify())


def GetName(soup: BeautifulSoup):
    return soup.find(id="etf-title").get_text(strip=True)  # type: ignore


def GetTickerAndCurrency(soup: BeautifulSoup):
    tickers = []

    h3_tags = soup.find_all("h3")
    for h3 in h3_tags:
        if h3.get_text(strip=True) == "Listings":
            div = h3.find_parent("div")

    table_body = div.find("tbody")
    rows = table_body.find_all("tr")
    for row in rows:
        divisions = row.find_all("td")
        if len(divisions) > 0:
            if divisions[0].get_text(strip=True) == "XETRA":
                currency = divisions[1].get_text(strip=True)
            ticker = divisions[2].get_text(strip=True)
            if ticker not in tickers and ticker != "-":
                tickers.append(ticker)
    return [tickers, currency]


def GetVolatility(soup: BeautifulSoup):
    rows = soup.find_all("tr")

    volatility = []

    for row in rows:
        divisions = row.find_all("td")
        if len(divisions) > 0:
            if divisions[0].get_text(strip=True) in [
                "Volatility 1 year",
                "Volatility 3 years",
                "Volatility 5 years",
            ]:
                volatility.append(divisions[1].get_text(strip=True))

    return volatility


def GetReturns(soup: BeautifulSoup):
    rows = soup.find_all("tr")

    volatility = []

    for row in rows:
        divisions = row.find_all("td")
        if len(divisions) > 0:
            if divisions[0].get_text(strip=True) in [
                "1 year",
                "3 years",
                "5 years",
            ]:
                volatility.append(divisions[1].get_text(strip=True))

    return volatility


def GetTotalExpenseRatio(soup: BeautifulSoup):
    rows = soup.find_all("tr")

    for row in rows:
        divisions = row.find_all("td")
        if len(divisions) > 0:
            if divisions[0].get_text(strip=True) == "Total expense ratio":
                ter = divisions[1].get_text(strip=True)
                break

    return ter


def GetDistribution(soup: BeautifulSoup):
    rows = soup.find_all("tr")

    for row in rows:
        divisions = row.find_all("td")
        if len(divisions) > 0:
            if divisions[0].get_text(strip=True) == "Distribution policy":
                distribution = divisions[1].get_text(strip=True)
                break

    return distribution


def GetReplication(soup: BeautifulSoup):
    rows = soup.find_all("tr")

    for row in rows:
        divisions = row.find_all("td")
        if len(divisions) > 0:
            if divisions[0].get_text(strip=True) == "Replication":
                distribution = divisions[1].get_text(strip=True)
                break

    return distribution


def GetCountries(soup: BeautifulSoup):
    countries = []

    h3_tags = soup.find_all("h3")
    for h3 in h3_tags:
        if h3.get_text(strip=True) == "Countries":
            div = h3.find_parent("div")

    table = div.find("table")
    rows = table.find_all("tr")

    for row in rows:
        divisions = row.find_all("td")
        if len(divisions) > 0:
            country = divisions[0].get_text(strip=True)
            percentage = divisions[1].get_text(strip=True)
            countries.append(country + " - " + percentage)

    return countries


def GetSectors(soup: BeautifulSoup):
    sectors = []

    h3_tags = soup.find_all("h3")
    for h3 in h3_tags:
        if h3.get_text(strip=True) == "Sectors":
            div = h3.find_parent("div")

    table = div.find("table")
    rows = table.find_all("tr")

    for row in rows:
        divisions = row.find_all("td")
        if len(divisions) > 0:
            sector = divisions[0].get_text(strip=True)
            percentage = divisions[1].get_text(strip=True)
            sectors.append(sector + " - " + percentage)

    return sectors


def GetFundSize(soup: BeautifulSoup):
    rows = soup.find_all("tr")

    for row in rows:
        divisions = row.find_all("td")
        if len(divisions) > 0:
            if divisions[0].get_text(strip=True) == "Fund size":
                fund_size = divisions[1].get_text(strip=True)
                break

    return fund_size


def GetNumberOfHoldings(soup: BeautifulSoup):
    divs = soup.find_all("div")
    for div in divs:
        if div.get_text(strip=True) == "Holdings":
            num_holdings = div.parent.find_all("div")[-1].get_text(
                strip=True
            )  # last div has the number of holdings
            break

    return num_holdings


def SeleniumScrape(list_isin: List[str]) -> Dict:
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.binary_location = BRAVE_PATH
    service = Service(CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)

    etf_data = {}

    for isin in list_isin:
        if not pd.isna(isin):
            etf_data[isin] = SeleniumScrapeETFs(driver, isin)

    # driver.quit()

    return etf_data


def SeleniumScrapeETFs(driver, isin: str):
    driver.get(URL_BASE + isin)

    try:
        element_present = EC.presence_of_element_located((By.ID, ID_TO_CHECK))
        WebDriverWait(driver, 10).until(element_present)
    except Exception as e:
        print(f"An error occurred: {e}")

    html = driver.page_source

    # with open("output.html", "r", encoding="utf-8") as file:
    #     html = file.read()

    soup = BeautifulSoup(html, "html.parser")

    etf_data = {}
    etf_data[DICT_NAME] = GetName(soup)
    etf_data[DICT_TICKERS], etf_data[DICT_CURRENCY] = GetTickerAndCurrency(soup)
    etf_data[DICT_VOLATILITY] = GetVolatility(soup)
    etf_data[DICT_RETURNS] = GetReturns(soup)
    etf_data[DICT_TER] = GetTotalExpenseRatio(soup)
    etf_data[DICT_DISTRIBUTION] = GetDistribution(soup)
    etf_data[DICT_REPLICATION] = GetReplication(soup)
    etf_data[DICT_COUNTRIES] = GetCountries(soup)
    etf_data[DICT_SECTORS] = GetSectors(soup)
    etf_data[DICT_FUND_SIZE] = GetFundSize(soup)
    etf_data[DICT_NUM_HOLDINGS] = GetNumberOfHoldings(soup)

    return etf_data


if __name__ == "__main__":
    SeleniumScrape(["IE00BK5BQT80"])
