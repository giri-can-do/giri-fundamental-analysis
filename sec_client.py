import requests


HEADERS = {
    "User-Agent": "Giri Fundamental Analysis Tool girishsangammanavar@gmail.com"
}


def get_company_cik(ticker):
    url = "https://www.sec.gov/files/company_tickers.json"

    response = requests.get(
        url,
        headers=HEADERS,
        timeout=30
    )

    response.raise_for_status()

    companies = response.json()

    for company in companies.values():

        if company["ticker"].upper() == ticker.upper():

            return str(
                company["cik_str"]
            ).zfill(10)

    return None


def get_company_facts(cik):
    url = (
        "https://data.sec.gov/api/xbrl/"
        f"companyfacts/CIK{cik}.json"
    )

    response = requests.get(
        url,
        headers=HEADERS,
        timeout=30
    )

    response.raise_for_status()

    return response.json()