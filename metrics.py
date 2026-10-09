METRICS = {
    "revenue": {
        "label": "Revenue",
        "concepts": [
            "RevenueFromContractWithCustomerExcludingAssessedTax",
            "Revenues",
            "SalesRevenueNet",
        ],
        "unit": "USD",
    },

    "gross_profit": {
        "label": "Gross Profit",
        "concepts": [
            "GrossProfit",
        ],
        "unit": "USD",
    },

    "operating_income": {
        "label": "Operating Income",
        "concepts": [
            "OperatingIncomeLoss",
        ],
        "unit": "USD",
    },

    "net_income": {
        "label": "Net Income",
        "concepts": [
            "NetIncomeLoss",
            "ProfitLoss",
        ],
        "unit": "USD",
    },

    "eps": {
        "label": "Diluted EPS",
        "concepts": [
            "EarningsPerShareDiluted",
        ],
        "unit": "USD/shares",
    },
}

BALANCE_SHEET_METRICS = {

    "cash": {
        "label": "Cash & Cash Equivalents",
        "concepts": [
            "CashAndCashEquivalentsAtCarryingValue",
            "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents",
        ],
        "unit": "USD",
    },

    "current_assets": {
        "label": "Current Assets",
        "concepts": [
            "AssetsCurrent",
        ],
        "unit": "USD",
    },

    "total_assets": {
        "label": "Total Assets",
        "concepts": [
            "Assets",
        ],
        "unit": "USD",
    },

    "current_liabilities": {
        "label": "Current Liabilities",
        "concepts": [
            "LiabilitiesCurrent",
        ],
        "unit": "USD",
    },

    "total_liabilities": {
        "label": "Total Liabilities",
        "concepts": [
            "Liabilities",
        ],
        "unit": "USD",
    },

    "equity": {
        "label": "Stockholders' Equity",
        "concepts": [
            "StockholdersEquity",
            "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
        ],
        "unit": "USD",
    },
}

CASH_FLOW_METRICS = {

    "operating_cash_flow": {
        "label": "Operating Cash Flow",
        "concepts": [
            "NetCashProvidedByUsedInOperatingActivities",
        ],
        "unit": "USD",
    },

    "capex": {
        "label": "Capital Expenditures",
        "concepts": [
            "PaymentsToAcquirePropertyPlantAndEquipment",
            "PaymentsForPropertyPlantAndEquipment",
        ],
        "unit": "USD",
    },
}