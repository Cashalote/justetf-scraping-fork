from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup

CHROMEDRIVER_PATH = "C:\\Users\\Hugo\\Documents\\chromedriver-win64\\chromedriver.exe"
BRAVE_PATH = "C:/Program Files/BraveSoftware/Brave-Browser/Application/brave.exe"
URL_BASE = "https://www.justetf.com/en/etf-profile.html?isin="
ID_TO_CHECK = "etf-title"


def SoupPrettyPrint(soup: BeautifulSoup):
    print(soup.prettify())


def GetName(soup: BeautifulSoup):
    return soup.find(id="etf-title").get_text(strip=True)


def GetTickerAndCurrency(soup: BeautifulSoup):
    rows = soup.find_all("tr")

    # Loop through rows to find "XETRA" and get the following values
    for row in rows:
        divisions = row.find_all("td")
        if len(divisions) > 0 and divisions[0].get_text(strip=True) == "XETRA":
            # Extract the text from each cell in the row following the "XETRA" cell
            currency = divisions[1].get_text(strip=True)
            ticker = divisions[2].get_text(strip=True)
            break

    return [ticker, currency]


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
    countries = {}

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
            countries[country] = percentage

    return countries


def GetSectors(soup: BeautifulSoup):
    sectors = {}

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
            sectors[sector] = percentage

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


# Change isin to be a list of isins read from excel
def SeleniumScrape(isin: str):
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.binary_location = BRAVE_PATH
    service = Service(CHROMEDRIVER_PATH)
    driver = webdriver.Chrome(service=service, options=options)

    SeleniumScrapeETFs(driver, isin)

    driver.quit()


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
    print(GetName(soup))
    print(GetTickerAndCurrency(soup))
    print(GetVolatility(soup))
    print(GetTotalExpenseRatio(soup))
    print(GetDistribution(soup))
    print(GetReplication(soup))
    print(GetCountries(soup))
    print(GetSectors(soup))
    print(GetFundSize(soup))
    print(GetNumberOfHoldings(soup))


if __name__ == "__main__":
    SeleniumScrape("IE00BK5BQT80")
