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

def get_hostname_with_nslookup(ip_address):
    try:
        result = subprocess.check_output(['nslookup', ip_address], stderr=subprocess.STDOUT, universal_newlines=True)
        lines = result.split('\n')
        for line in lines:
            if 'Name:' in line:
                hostname = line.split('Name:')[1].strip()
                return hostname
        return None
    except subprocess.CalledProcessError:
        return None

def get_hostname_with_ping(ip_address):
    try:
        result = subprocess.check_output(['ping', '-a', ip_address], stderr=subprocess.STDOUT, universal_newlines=True)
        hostname = result.split(' ')[1].strip('[]')
        return hostname
    except subprocess.CalledProcessError:
        return None

def detect_os(hostname):
    try:
        result = subprocess.run(['ping', '-n', '1', hostname], capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            if "Reply" in result.stdout:
                return "Windows"
            else:
                return "Linux"
    except Exception as e:
        print(f"Ошибка при определении операционной системы для хостнейма {hostname}: {e}")
    return "Unknown"

def add_os_info_to_sheet(sheet):
    max_row = sheet.max_row
    os_sheet = wb.create_sheet(title="OS Info")
    os_sheet.append(["Hostname", "OS"])
    for row in range(2, max_row + 1):
        hostname = sheet.cell(row=row, column=3).value
        if not hostname:
            hostname = sheet.cell(row=row, column=5).value
        if hostname:
            detected_os = detect_os(hostname)
            os_sheet.append([hostname, detected_os])

def process_hostnames(sheet):
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
                hostname = get_hostname_with_nslookup(ip_address)
                if not hostname:
                    hostname = get_hostname_with_ping(ip_address)
                if hostname:
                    sheet.cell(row=row, column=3).value = hostname
                else:
                    sheet.cell(row=row, column=3).value = "Не удалось получить хостнейм"
            except Exception as e:
                print(f"Ошибка при обработке хостнейма с IP {ip_address}: {e}")

# Открываем файл hosts.xlsx
wb = openpyxl.load_workbook('hosts.xlsx')
sheet = wb.active

# Обработка хостнеймов
process_hostnames(sheet)

# Добавляем информацию об операционных системах
add_os_info_to_sheet(sheet)

# Сохраняем изменения
wb.save('hosts.xlsx')

print("Готово!")
