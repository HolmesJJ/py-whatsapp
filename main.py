import os
import pywhatkit
import pandas as pd


def test():
    pywhatkit.sendwhatmsg_instantly('+65', 'Hi, how are you?', 15, True, 2)
    pywhatkit.sendwhatmsg_instantly('+65', 'Hi, how are you?', 15, True, 2)


def run():
    folder_path = os.path.dirname(os.path.abspath(__file__))
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
                        pywhatkit.sendwhatmsg_instantly(f'+65{col2}', col10, 10, True, 2)
                        print('Phone:', col2, 'Message:', col10)
            except Exception as e:
                print(f'Read {filename} error: {e}')


if __name__ == '__main__':
    run()
