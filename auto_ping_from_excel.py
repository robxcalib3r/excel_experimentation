from openpyxl import load_workbook
import os
import platform
import subprocess

def ping_ip(ip):
    # Option for the number of packers as a function of OS
    param = '-n' if platform.system().lower() == 'windows' else '-c'
    command = ['ping', param, '1', ip]
    # response = os.system(f"ping -c 3 {ip}")
    return os.system(' '.join(command)) == 0

def ping_ip_subprocess_Exists(ip):
    # Option for the number of packers as a function of OS
    param = '-n' if platform.system().lower() == 'windows' else '-c'
    command = ['ping', param, '3', ip]
    response = subprocess.run(command, capture_output=True)
    return response.returncode == 0

def modifyExcelByPing(file_path, read_column, write_column):
    workbook = load_workbook(filename=file_path)

    sheet = workbook.active
    for row in range(2, sheet.max_row):
        ip = sheet.cell(row=row, column=read_column).value
        if ip is not None:
            status = 'online' if ping_ip_subprocess_Exists(ip) else 'offline'
            print(f'{ip} : {status}')
            sheet.cell(row=row, column=write_column, value=status)
        else:
            status = 'Not valid IP'
            print(status)
            sheet.cell(row=row, column=write_column, value=status)
    workbook.save(file_path)

if __name__ == "__main__":
    modifyExcelByPing("all_veam_backup_list_6-22.xlsm", 3, 4)