import re
import os
import shutil
import telegram
import database
import pandas as pd
from datetime import datetime, timedelta, date
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook


def data_xlsm():
    read_not = False
    today = date.today()
    dir_path = r'\powerbi\UPLOAD_FILE'
    backup_path = r'\powerbi\Backup'
    failed_path = r'\powerbi\Failed'
    for path in os.listdir(dir_path):
        if os.path.isfile(os.path.join(dir_path, path)):
            read_not = re.search("world_assets", path)
            if read_not:
                dir_path = f'\powerbi\UPLOAD_FILE\{path}'
                break
    ###
    if read_not:
        list_data = []
        wb = load_workbook(filename=dir_path)
        sheet = wb.active
        data_date = sheet.cell(row=2, column=2).value
        date_format = data_date.strftime("%m/%d/%Y")
        if date_format == today.strftime("%m/%d/%Y"):
            for row_index in range(2, sheet.max_row + 1):
                data = {}
                for column_index in range(1, sheet.max_column + 1):
                    column_letter = get_column_letter(column_index)
                    if row_index > 1:
                        key = sheet[column_letter + str(1)].value
                        value = sheet[column_letter + str(row_index)].value
                        data[key] = value
                list_data.append(data)
            list_data_success = []
            sub_day = 0
            for x in range(0, 200 + 1):
                arr = []
                for data in list_data:
                    if f'DongTien T-{x}' or f'DongTien T{x}' in data:
                        obj = {}
                        date_time = data['Date/Time'] - timedelta(days=sub_day)
                        th_day = pd.Timestamp(date_time).dayofweek
                        if th_day == 6:
                            sub_day = sub_day + 2
                        date_time = data['Date/Time'] - timedelta(days=sub_day)
                        obj['asset_code'] = data['Ticker']
                        obj['datetime'] = date_time.strftime("%Y-%m-%d %H:%M:%S")
                        if x == 0:
                            obj['cashflow'] = data[f'DongTien T{x}']
                        else:
                            obj['cashflow'] = data[f'DongTien T-{x}']
                        arr.append(obj)
                sub_day = sub_day + 1
                arr = sorted(arr, key=lambda k: k['cashflow'], reverse=False)
                for n in range(0, len(arr)):
                    arr[n]['top'] = n + 1
                    arr[n]['ranks'] = arr[n]['top'] / len(arr) * 10
                    list_data_success.append(arr[n])

            # for get in list_data_success:
            cur = database.main.cursor()
            cur.executemany('''INSERT INTO datafeed.dsc_mst_asset_cashflow_555 (asset_code, datetime, cashflow, ranks)
            VALUES (%s, %s, %s, %s)ON CONFLICT ON CONSTRAINT dsc_mst_asset_cashflow_pkey_555 
            DO UPDATE SET (asset_code, datetime, cashflow, ranks) = 
            (EXCLUDED.asset_code, EXCLUDED.datetime, EXCLUDED.cashflow, EXCLUDED.ranks)''',
                            [[item.get("asset_code"),
                              item.get("datetime"),
                              item.get("cashflow"),
                              item.get("ranks")] for item in list_data_success])
            database.main.commit()
            print(f'[world_assets] insert done data size: {len(list_data_success)}', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            print(f'[world_assets] Data ngày: {data_date} - 200')
            telegram.send(mess=f"[world_assets] Insert done / Data size: {len(list_data_success)}")
            shutil.move(dir_path, backup_path)
        else:
            shutil.move(dir_path, failed_path)
            print(f'[transaction_stocks] Không phải data ngày: {today}. Đây là data ngày: {data_date}')
            telegram.send(mess=f"[transaction_stocks] Sai data. Đây là data ngày: {data_date}")
        read_not = False
    else:
        print("[world_assets] Không thấy file !!!")
        telegram.send(mess=f"[world_assets] Không tìm thấy file ngày: {today}")