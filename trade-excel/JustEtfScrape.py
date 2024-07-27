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

    # Loop through rows to find "XETRA" and get the following values
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

    # Loop through rows to find "XETRA" and get the following values
    for row in rows:
        divisions = row.find_all("td")
        if len(divisions) > 0:
            if divisions[0].get_text(strip=True) == "Total expense ratio":
                div = divisions[1].find("div")
                ter = div.get_text(strip=True)
                break

    return ter


# Change isin to be a list of isins
def SeleniumScrape(isin: str):
    # options = webdriver.ChromeOptions()
    # options.add_argument("--headless")
    # options.binary_location = BRAVE_PATH
    # service = Service(CHROMEDRIVER_PATH)
    # driver = webdriver.Chrome(service=service, options=options)

    # SeleniumScrapeETF(driver, isin)

    # driver.quit()

    # html = driver.page_source

    with open("output.html", "r", encoding="utf-8") as file:
        html_content = file.read()

    soup = BeautifulSoup(html_content, "html.parser")
    print(GetTickerAndCurrency(soup))
    print(GetVolatility(soup))
    print(GetTotalExpenseRatio(soup))


def SeleniumScrapeETF(driver, isin: str):
    driver.get(URL_BASE + isin)

    try:
        element_present = EC.presence_of_element_located((By.ID, ID_TO_CHECK))
        WebDriverWait(driver, 10).until(element_present)
    except Exception as e:
        print(f"An error occurred: {e}")


if __name__ == "__main__":
    SeleniumScrape("IE00BK5BQT80")
