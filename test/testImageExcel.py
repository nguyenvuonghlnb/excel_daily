import openpyxl
import os
import re

readornotread = False
dir_path = r'\\192.168.50.12\Public\CNTT\Test\File'
backup_path = r'\\192.168.50.12\Public\CNTT\Test\Backup'
failed_path = r'\\192.168.50.12\Public\CNTT\Test\Failed'
dir_path_file = ''
for path in os.listdir(dir_path):
    if os.path.isfile(os.path.join(dir_path, path)):
        readornotread = re.search("checker", path)
        if readornotread:
            dir_path_file = f'\\\\192.168.50.12\\Public\\CNTT\\Test\\File\\{path}'
pxl_doc = openpyxl.load_workbook(dir_path_file)
sheet = pxl_doc.active
pxl_doc.save(f'{dir_path}\\checker.xlsx')
os.remove(dir_path_file)




