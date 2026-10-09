import sys

from sec_client import (
    get_company_cik,
    get_company_facts,
)

from analyzer import (
    build_income_statement,
    build_profitability_analysis,
    build_balance_sheet,
    build_financial_health_analysis,
    build_cash_flow_statement,
    build_cash_flow_analysis,
)
from excel_exporter import create_workbook

def print_income_statement(statement):

    print("\n" + "=" * 80)
    print("5-YEAR INCOME STATEMENT")
    print("=" * 80)

    for metric in statement.values():

        print(f"\n{metric['label']}")
        print("-" * 50)

        previous_value = None

        for record in metric["records"]:

            value = record["val"]

            if metric["unit"] == "USD":
                display = (
                    f"${value / 1_000_000_000:,.2f}B"
                )
            else:
                display = f"${value:,.2f}"

            if previous_value is None:
                growth = "N/A"

            elif previous_value != 0:

                growth_rate = (
                    (value - previous_value)
                    / abs(previous_value)
                    * 100
                )

                growth = f"{growth_rate:+.2f}%"

            else:
                growth = "N/A"

            print(
                f"{record['end']} | "
                f"{display:<15} | "
                f"Growth: {growth}"
            )

            previous_value = value


def print_profitability_analysis(analysis):

    print("\n" + "=" * 80)
    print("PROFITABILITY ANALYSIS")
    print("=" * 80)

    print(
        f"{'Year':<15}"
        f"{'Gross Margin':<20}"
        f"{'Operating Margin':<20}"
        f"{'Net Margin':<20}"
    )

    print("-" * 75)

    for row in analysis:

        gross = (
            f"{row['gross_margin']:.2f}%"
            if row["gross_margin"] is not None
            else "N/A"
        )

        operating = (
            f"{row['operating_margin']:.2f}%"
            if row["operating_margin"] is not None
            else "N/A"
        )

        net = (
            f"{row['net_margin']:.2f}%"
            if row["net_margin"] is not None
            else "N/A"
        )

        print(
            f"{row['period_end']:<15}"
            f"{gross:<20}"
            f"{operating:<20}"
            f"{net:<20}"
        )

def print_balance_sheet(statement):

    print("\n" + "=" * 80)
    print("5-YEAR BALANCE SHEET")
    print("=" * 80)

    for metric in statement.values():

        print(f"\n{metric['label']}")
        print("-" * 50)

        for record in metric["records"]:

            value = record["val"]

            display = (
                f"${value / 1_000_000_000:,.2f}B"
            )

            print(
                f"{record['end']} | "
                f"{display}"
            )

def print_financial_health_analysis(analysis):

    print("\n" + "=" * 80)
    print("FINANCIAL HEALTH ANALYSIS")
    print("=" * 80)

    print(
        f"{'Year':<15}"
        f"{'Current Ratio':<20}"
        f"{'Liab / Assets':<20}"
        f"{'Equity Ratio':<20}"
    )

    print("-" * 75)

    for row in analysis:

        current_ratio = (
            f"{row['current_ratio']:.2f}"
            if row["current_ratio"] is not None
            else "N/A"
        )

        liabilities = (
            f"{row['liabilities_to_assets']:.2f}%"
            if row["liabilities_to_assets"] is not None
            else "N/A"
        )

        equity = (
            f"{row['equity_ratio']:.2f}%"
            if row["equity_ratio"] is not None
            else "N/A"
        )

        print(
            f"{row['period_end']:<15}"
            f"{current_ratio:<20}"
            f"{liabilities:<20}"
            f"{equity:<20}"
        )

def print_cash_flow_analysis(analysis):

    print("\n" + "=" * 95)
    print("CASH FLOW ANALYSIS")
    print("=" * 95)

    print(
        f"{'Year':<15}"
        f"{'Operating CF':<18}"
        f"{'CapEx':<15}"
        f"{'Free CF':<18}"
        f"{'FCF Margin':<15}"
        f"{'Cash Conv.':<15}"
    )

    print("-" * 95)

    for row in analysis:

        ocf = (
            f"${row['operating_cash_flow'] / 1_000_000_000:.2f}B"
        )

        capex = (
            f"${row['capex'] / 1_000_000_000:.2f}B"
            if row["capex"] is not None
            else "N/A"
        )

        fcf = (
            f"${row['free_cash_flow'] / 1_000_000_000:.2f}B"
            if row["free_cash_flow"] is not None
            else "N/A"
        )

        fcf_margin = (
            f"{row['fcf_margin']:.2f}%"
            if row["fcf_margin"] is not None
            else "N/A"
        )

        cash_conversion = (
            f"{row['cash_conversion']:.2f}%"
            if row["cash_conversion"] is not None
            else "N/A"
        )

        print(
            f"{row['period_end']:<15}"
            f"{ocf:<18}"
            f"{capex:<15}"
            f"{fcf:<18}"
            f"{fcf_margin:<15}"
            f"{cash_conversion:<15}"
        )


def main():

    if len(sys.argv) < 2:
        print(
            "Usage: python app.py TICKER"
        )
        return

    ticker = sys.argv[1].upper()

    print(f"\nAnalyzing {ticker}...")

    cik = get_company_cik(ticker)

    if not cik:
        print(
            f"Ticker '{ticker}' was not found."
        )
        return

    company_facts = get_company_facts(cik)

    company_name = company_facts["entityName"]

    print(f"Company: {company_name}")
    print(f"Ticker:  {ticker}")
    print(f"CIK:     {cik}")

    income_statement = build_income_statement(
        company_facts,
        years=5
    )

    profitability = build_profitability_analysis(
        income_statement
    )

    print_income_statement(
        income_statement
    )

    print_profitability_analysis(
        profitability
    )

    balance_sheet = build_balance_sheet(
        company_facts,
        years=5
    )

    print_balance_sheet(
        balance_sheet
    )


    financial_health_analysis = build_financial_health_analysis(balance_sheet)

    print_financial_health_analysis(
        financial_health_analysis
    )

    cash_flow_statement = build_cash_flow_statement(
        company_facts,
        years=5
    )

    cash_flow_analysis = build_cash_flow_analysis(
        cash_flow_statement,
        income_statement
    )

    print_cash_flow_analysis(
        cash_flow_analysis
    )

    excel_file = create_workbook(
        ticker=ticker,
        company_name=company_name,
        income_statement=income_statement,
        profitability=profitability,
        balance_sheet=balance_sheet,
        financial_health=financial_health_analysis,
        cash_flow_analysis=cash_flow_analysis,
    )

    print(
        f"\nExcel workbook created:"
        f"\n{excel_file}"
    )

if __name__ == "__main__":
    main()