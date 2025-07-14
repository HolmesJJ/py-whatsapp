import os
import sys
import time
import argparse
import pywhatkit

from openpyxl import load_workbook


DEFAULT_WAIT = 20


def test():
    pywhatkit.sendwhatmsg_instantly('+65', 'Hi, how are you?', DEFAULT_WAIT, True)
    pywhatkit.sendwhatmsg_instantly('+65', 'Hi, how are you?', DEFAULT_WAIT, True)


def get_folder_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def run(wait=DEFAULT_WAIT):
    print('--- Start ---')
    folder_path = get_folder_path()
    print(f'--- Working in: {folder_path} ---')
    error = False
    for file_name in os.listdir(folder_path):
        if file_name.lower().endswith('.xlsx'):
            file_path = os.path.join(folder_path, file_name)
            try:
                print(f'--- Processing Excel {file_name} ... ---')
                wb = load_workbook(file_path, data_only=True)
                ws = wb.active
                header = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
                try:
                    idx_id = header.index('ID') + 1
                    idx_mobile = header.index('Mobile Number') + 1
                    idx_name = header.index('Official Name') + 1
                    idx_message = header.index('Full Text Message') + 1
                    idx_sent = header.index('Sent') + 1
                except ValueError as e2:
                    print(f'! Missing columns: {e2}')
                    wb.close()
                    error = True
                    continue
                for row in ws.iter_rows(min_row=2):
                    id_cell = row[idx_id - 1]
                    mobile_cell = row[idx_mobile - 1]
                    name_cell = row[idx_name - 1]
                    message_cell = row[idx_message - 1]
                    sent_cell = row[idx_sent - 1]
                    if not mobile_cell.value:
                        print(f'! Row {id_cell.value} missing mobile value')
                        wb.close()
                        error = True
                        break
                    if not name_cell.value:
                        print(f'! Row {id_cell.value} missing name value')
                        wb.close()
                        error = True
                        break
                    if not message_cell.value:
                        print(f'! Row {id_cell.value} missing message value')
                        wb.close()
                        error = True
                        break
                    if sent_cell.value == 1:
                        continue
                    mobile = str(mobile_cell.value).strip()
                    message = str(message_cell.value).strip()
                    name = str(name_cell.value).strip()
                    full_number = f'+65{mobile}'
                    print(f'→ {name} | {full_number} | wait={wait}s | message={message[:30]}...')
                    try:
                        pywhatkit.sendwhatmsg_instantly(f'+{full_number}', message, wait, True)
                        sent_cell.value = 1
                        print(f'! Updating Excel file: {file_name}. Please DO NOT close or terminate the program...')
                        wb.save(file_path)
                        print(f'! Excel file {file_name} has been updated. '
                              f'You may now safely close or terminate the program...')
                        print('→ Waiting 3 seconds...', end='', flush=True)
                        for i in range(3, 0, -1):
                            print(f' {i}...', end='', flush=True)
                            time.sleep(1)
                        print()
                    except Exception as e3:
                        print(f'! Send failed: {e3}')
                wb.close()
            except Exception as e1:
                print(f'! Read {file_name} error: {e1}')
                error = True
            if error:
                break
    print('--- End ---')


def parse_args():
    parser = argparse.ArgumentParser(description='Bulk-send WhatsApp messages based on Excel files.')
    parser.add_argument('--wait', type=int, default=DEFAULT_WAIT,
                        help=f'Wait (seconds) before sending. The default is {DEFAULT_WAIT} seconds.')
    return parser.parse_args()


if __name__ == '__main__':
    # pyinstaller --onefile --console --hidden-import=pywhatkit --hidden-import=openpyxl main.py
    args = parse_args()
    run(args.wait)
