from FinanceBot import FinanceBot


def main():
    #Filename, start_row, start_col
    finance_bot = FinanceBot("HK-Stock-Analysis - Copy", 56, 2)
    # finance_bot.initialize_worksheet()
    finance_bot.extract_finance_data()

if __name__ == "__main__":
    # format_worksheet()
    main()
