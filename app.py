from FinanceBot import FinanceBot


def main():
    finance_bot = FinanceBot("copy_hk_stock")
    finance_bot.extract_finance_data()


if __name__ == "__main__":
    # format_worksheet()
    main()
