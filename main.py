import time
import config
from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler
from controller import dsc_mst_asset_cashflow
from controller import dsc_asset_assess
from controller import dsc_industry_growth
from controller import dsc_stock_transaction_net
from controller import dsc_market_transaction_net

scheduler = BlockingScheduler()
print("Start deploying !", datetime.now().strftime('%Y-%m-%d %H:%M:%S'))


def import_excel_daily():
    date_request = datetime.now().strftime('%Y-%m-%d')
    # xu hướng tài sản (daily)
    dsc_mst_asset_cashflow.data_xlsm()

    # backtest
    dsc_asset_assess.data_xlsx_backtest()

    # industry_growth_lv2
    dsc_industry_growth.data_xlsm_growth()

    # transaction_stocks (daily)
    dsc_stock_transaction_net.data_xlsm_stock_transaction()

    # transaction_day (daily)
    dsc_market_transaction_net.data_xlsm_market_transaction()


# if __name__ == '__main__':
#     import_excel_daily()

scheduler.add_job(import_excel_daily, 'cron', day_of_week='mon-sun', hour='11', minute='13', start_date='2022-09-06 08:00:00', end_date='2023-09-06 08:00:00', timezone='Asia/Ho_Chi_Minh')
# Start the scheduler
scheduler.start()


