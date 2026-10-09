from metrics import (
    METRICS,
    BALANCE_SHEET_METRICS,
    CASH_FLOW_METRICS,
)


def find_metric_concept(company_facts, metric_config):

    us_gaap = company_facts["facts"]["us-gaap"]

    for concept in metric_config["concepts"]:

        if concept in us_gaap:
            return concept

    return None


def get_annual_metric_by_config(
    company_facts,
    metric_config,
    years=5
):

    concept = find_metric_concept(
        company_facts,
        metric_config
    )

    if not concept:
        return []

    us_gaap = company_facts["facts"]["us-gaap"]

    metric = us_gaap[concept]

    unit = metric_config["unit"]

    records = metric["units"].get(unit, [])

    annual_records = []

    for record in records:

        if (
            record.get("form") == "10-K"
            and record.get("fp") == "FY"
        ):
            annual_records.append(record)

    # Remove duplicate comparative values
    by_period = {}

    for record in annual_records:

        period_end = record["end"]

        if (
            period_end not in by_period
            or record["filed"]
            > by_period[period_end]["filed"]
        ):
            by_period[period_end] = record

    clean_records = list(
        by_period.values()
    )

    clean_records.sort(
        key=lambda x: x["end"]
    )

    return clean_records[-years:]


def build_income_statement(
    company_facts,
    years=5
):

    statement = {}

    for metric_key, config in METRICS.items():

        records = get_annual_metric_by_config(
            company_facts,
            config,
            years
        )

        statement[metric_key] = {
            "label": config["label"],
            "unit": config["unit"],
            "records": records,
        }

    return statement

def build_profitability_analysis(
    income_statement
):

    revenue = {
        record["end"]: record["val"]
        for record
        in income_statement["revenue"]["records"]
    }

    gross_profit = {
        record["end"]: record["val"]
        for record
        in income_statement["gross_profit"]["records"]
    }

    operating_income = {
        record["end"]: record["val"]
        for record
        in income_statement["operating_income"]["records"]
    }

    net_income = {
        record["end"]: record["val"]
        for record
        in income_statement["net_income"]["records"]
    }

    analysis = []

    for period_end, revenue_value in revenue.items():

        if revenue_value == 0:
            continue

        analysis.append({
            "period_end": period_end,

            "gross_margin":
                gross_profit.get(period_end)
                / revenue_value * 100
                if period_end in gross_profit
                else None,

            "operating_margin":
                operating_income.get(period_end)
                / revenue_value * 100
                if period_end in operating_income
                else None,

            "net_margin":
                net_income.get(period_end)
                / revenue_value * 100
                if period_end in net_income
                else None,
        })

    return analysis

def get_annual_balance_metric_by_config(
    company_facts,
    metric_config,
    years=5
):
    concept = find_metric_concept(
        company_facts,
        metric_config
    )

    if not concept:
        return []

    us_gaap = company_facts["facts"]["us-gaap"]
    metric = us_gaap[concept]

    unit = metric_config["unit"]
    records = metric["units"].get(unit, [])

    annual_records = []

    for record in records:

        if (
            record.get("form") == "10-K"
            and record.get("fp") == "FY"
        ):
            annual_records.append(record)

    # Balance sheet facts represent a point in time,
    # so use the END date as the reporting date.
    by_period = {}

    for record in annual_records:

        period_end = record["end"]

        if (
            period_end not in by_period
            or record["filed"] > by_period[period_end]["filed"]
        ):
            by_period[period_end] = record

    clean_records = list(by_period.values())

    clean_records.sort(
        key=lambda x: x["end"]
    )

    return clean_records[-years:]

def build_balance_sheet(
    company_facts,
    years=5
):
    statement = {}

    for metric_key, config in BALANCE_SHEET_METRICS.items():

        records = get_annual_balance_metric_by_config(
            company_facts,
            config,
            years
        )

        statement[metric_key] = {
            "label": config["label"],
            "unit": config["unit"],
            "records": records,
        }

    return statement

def build_financial_health_analysis(balance_sheet):

    def to_period_dict(metric_key):
        return {
            record["end"]: record["val"]
            for record
            in balance_sheet[metric_key]["records"]
        }

    current_assets = to_period_dict(
        "current_assets"
    )

    current_liabilities = to_period_dict(
        "current_liabilities"
    )

    total_assets = to_period_dict(
        "total_assets"
    )

    total_liabilities = to_period_dict(
        "total_liabilities"
    )

    equity = to_period_dict(
        "equity"
    )

    analysis = []

    for period_end, assets in total_assets.items():

        row = {
            "period_end": period_end,
            "current_ratio": None,
            "liabilities_to_assets": None,
            "equity_ratio": None,
        }

        if (
            period_end in current_assets
            and period_end in current_liabilities
            and current_liabilities[period_end] != 0
        ):
            row["current_ratio"] = (
                current_assets[period_end]
                / current_liabilities[period_end]
            )

        if (
            period_end in total_liabilities
            and assets != 0
        ):
            row["liabilities_to_assets"] = (
                total_liabilities[period_end]
                / assets
            ) * 100

        if (
            period_end in equity
            and assets != 0
        ):
            row["equity_ratio"] = (
                equity[period_end]
                / assets
            ) * 100

        analysis.append(row)

    return analysis

def build_cash_flow_statement(
    company_facts,
    years=5
):
    statement = {}

    for metric_key, config in CASH_FLOW_METRICS.items():

        records = get_annual_metric_by_config(
            company_facts,
            config,
            years
        )

        statement[metric_key] = {
            "label": config["label"],
            "unit": config["unit"],
            "records": records,
        }

    return statement

def build_cash_flow_analysis(
    cash_flow_statement,
    income_statement
):

    operating_cash_flow = {
        record["end"]: record["val"]
        for record
        in cash_flow_statement[
            "operating_cash_flow"
        ]["records"]
    }

    capex = {
        record["end"]: record["val"]
        for record
        in cash_flow_statement[
            "capex"
        ]["records"]
    }

    revenue = {
        record["end"]: record["val"]
        for record
        in income_statement[
            "revenue"
        ]["records"]
    }

    net_income = {
        record["end"]: record["val"]
        for record
        in income_statement[
            "net_income"
        ]["records"]
    }

    analysis = []

    for period_end, ocf in operating_cash_flow.items():

        row = {
            "period_end": period_end,
            "operating_cash_flow": ocf,
            "capex": None,
            "free_cash_flow": None,
            "fcf_margin": None,
            "cash_conversion": None,
        }

        if period_end in capex:

            capex_value = capex[period_end]

            free_cash_flow = (
                ocf - capex_value
            )

            row["capex"] = capex_value
            row["free_cash_flow"] = free_cash_flow

            if (
                period_end in revenue
                and revenue[period_end] != 0
            ):
                row["fcf_margin"] = (
                    free_cash_flow
                    / revenue[period_end]
                    * 100
                )

        if (
            period_end in net_income
            and net_income[period_end] != 0
        ):
            row["cash_conversion"] = (
                ocf
                / net_income[period_end]
                * 100
            )

        analysis.append(row)

    return analysis