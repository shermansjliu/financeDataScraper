from enum import Enum

class Info(Enum):
    COMPANY_NAME = "Company Name"
    EVTS ="Enterprise Value to Sales"
    EBIT = "EBIT"
    S_R = "Sales/Revenue"
    CASH_ST_INVESTMENTS = "Cash & Short Term Investments" #Cash Only on wsj
    TCA = "Total Current Assets"
    STDebt = "Short Term Debt"
    CPLTDebt = "Current Portion of Long Term Debt"
    LTDebt = "Long-Term Debt"
    TOTDebt = "Total Debt"
    TCL = "Total Current Liabilities"
    NETPPEQ = "Net Property, Plant & Equipment"