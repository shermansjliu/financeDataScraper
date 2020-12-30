from FinanceBot import FinanceBot


def main():
    finance_bot = FinanceBot("HK-Stock-Analysis - Copy")
    finance_bot.initialize_worksheet()
    finance_bot.extract_finance_data()


if __name__ == "__main__":
    # format_worksheet()
    main()
