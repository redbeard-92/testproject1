import openpyxl
import subprocess
import socket

def check_ping(ip_address):
    try:
        result = subprocess.run(['ping', '-n', '1', ip_address], capture_output=True, text=True, timeout=5)
        return result.returncode == 0
    except Exception as e:
        print(f"Ошибка при выполнении ping для {ip_address}: {e}")
        return False
def get_hostname(ip_address):
    try:
        result = subprocess.check_output(['ping', '-a', ip_address], stderr=subprocess.STDOUT, universal_newlines=True)
        hostname = result.split(' ')[1].strip('[]')
        return hostname
    except subprocess.CalledProcessError:
        return None

def detect_os(ip_address):
    try:
        result = subprocess.run(['ping', '-n', '1', ip_address], capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            if "Reply" in result.stdout:
                return "Windows"
            else:
                return "Linux"
    except Exception as e:
        print(f"Ошибка при определении операционной системы для IP {ip_address}: {e}")
    return "Unknown"

def process_devices(sheet):
    max_row = sheet.max_row
    for row in range(2, max_row + 1):
        ip_address = sheet.cell(row=row, column=2).value
        if ip_address:
            print(f"Пингуется IP: {ip_address}")
            try:
                if check_ping(ip_address):
                    sheet.cell(row=row, column=4).value = "Доступен"
                else:
                    sheet.cell(row=row, column=4).value = "Недоступен"

                hostname = get_hostname(ip_address)
                if hostname:
                    sheet.cell(row=row, column=3).value = hostname.encode('ascii', 'ignore').decode('ascii')
                else:
                    sheet.cell(row=row, column=3).value = "Не удалось получить хостнейм"

                os = detect_os(ip_address)
                sheet.cell(row=row, column=5).value = os
            except Exception as e:
                print(f"Ошибка при обработке устройства с IP {ip_address}: {e}")

def process_hostnames(sheet):
    max_row = sheet.max_row
    for row in range(2, max_row + 1):
        hostname = sheet.cell(row=row, column=6).value
        if hostname:
            print(f"Пингуется хостнейм: {hostname}")
            try:
                ip_address = socket.gethostbyname(hostname)
                sheet.cell(row=row, column=7).value = ip_address
                if check_ping(ip_address):
                    sheet.cell(row=row, column=8).value = "Доступен"
                else:
                    sheet.cell(row=row, column=8).value = "Недоступен"
                os = detect_os(ip_address)
                sheet.cell(row=row, column=9).value = os
            except Exception as e:
                print(f"Ошибка при обработке хостнейма {hostname}: {e}")

wb = openpyxl.load_workbook('hosts.xlsx')
sheet = wb.active
process_devices(sheet)
process_hostnames(sheet)
wb.save('hosts.xlsx')
print("Готово!")
