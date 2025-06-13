import os
import sys
import pywhatkit
import pandas as pd


def test():
    pywhatkit.sendwhatmsg_instantly('+65', 'Hi, how are you?', 20, True, 2)
    pywhatkit.sendwhatmsg_instantly('+65', 'Hi, how are you?', 20, True, 2)


def get_folder_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def run():
    print('--- Start ---')
    folder_path = get_folder_path()
    print(f'--- {folder_path} ---')
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(folder_path, filename)
            try:
                print(f'--- Reading Excel {filename} ... ---')
                df = pd.read_excel(file_path, engine='openpyxl')
                print(df)
                print(f'--- Sending Messages ... ---')
                if df.shape[1] >= 8:
                    for index, row in df.iterrows():
                        col2 = row.iloc[4]
                        col10 = row.iloc[8]
                        print(f'---')
                        print(f'Sending message to phone {col2} in 20 seconds:')
                        pywhatkit.sendwhatmsg_instantly(f'+65{col2}', col10, 20, True, 2)
                        print('Message:', col10)
                        print(f'---')
            except Exception as e:
                print(f'Read {filename} error: {e}')
    print('--- End ---')


if __name__ == '__main__':
    # pyinstaller --onefile --console --hidden-import=pywhatkit --hidden-import=pandas --hidden-import=openpyxl main.py
    run()
