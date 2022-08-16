import time
import pandas as pd
import datetime
import win32com.client
from pythoncom import com_error
import numpy as np


def check_convert_str_float(df, column):
    if isinstance(df[column][0], str):
        df[column] = df[column].str.replace(',', '')
        df[column] = df[column].str.replace('[', '', regex=True)
        df[column] = df[column].replace(']', np.NaN, regex=True)
        df[column] = df[column].astype(float)
    return df



def Convert_data(filename, usecols):
    df = pd.read_csv(filename, header=None, usecols=usecols,
                     names = ['time', 'price', 'change', 'pct change'],
                     index_col=['time'], parse_dates=['time'])

    df = check_convert_str_float(df, 'price')
    df.fillna(method='ffill', inplace = True)

    data = df['price'].resample('1Min').ohlc()

    data['time'] = data.index
    data['time'] = pd.to_datetime(data['time'], format = "%Y-%m-%d %H:%M:%S")
    data = data[['time', 'open', 'high', 'low', 'close']]

    data.reset_index(drop=True, inplace = True)
    return data


# Step 2
def Process_Data():
    df=pd.read_excel(path)
    data = df.iloc[:, 1:4].values.tolist()
    data = [round(item, 6) for sublist in data for item in sublist]

    time_stamp = datetime.datetime.now() - datetime.timedelta(hours=12)
    time_stamp = time_stamp.strftime("%Y-%m-%d %H:%M:%S")
    data.insert(0, time_stamp)

    df_data = pd.DataFrame(data)
    df_data = df_data.T
    df_data.to_csv('Raw Data.csv',mode='a', header=False, index=False)

    data_ohlc_F = Convert_data('Raw Data.csv', [0, 1, 2, 3])
    data_ohlc_Am = Convert_data('Raw Data.csv', [0, 4, 5, 6])
    data_ohlc_A = Convert_data('Raw Data.csv', [0, 7, 8, 9])
    data_ohlc_N = Convert_data('Raw Data.csv', [0, 10, 11, 12])
    data_ohlc_G = Convert_data('Raw Data.csv', [0, 13, 14, 15])
    result = pd.concat([data_ohlc_F, data_ohlc_Am.iloc[:, 1:5], data_ohlc_A.iloc[:, 1:5],
                        data_ohlc_N.iloc[:, 1:5], data_ohlc_G.iloc[:, 1:5]], axis = 1, join="inner")

    result.to_csv('Raw Data_ohlc.csv', mode='w', header=False, index=False)

    

    print(data)


# Step 1
def Refresh_Save():
    #1 Connect this to an excel and workbook
    xlApp = win32com.client.DispatchEx("Excel.Application")  

    try:
        Workbook = xlApp.Workbooks.Open(path)
        ready_to_open = True
    except(AttributeError, com_error):  
        print("Workbook Open Error")
        ready_to_open = False


    if ready_to_open:

        ready_to_display = True
        while ready_to_display:
            try:
                xlApp.DisplayAlerts = False   
                ready_to_display = False
            except (AttributeError, com_error):
                time.sleep(1)
                print('Workbook Display Error')
                ready_to_display = True

        #2 refresh that workbook (1s/2s)
        ready_to_refresh = True
        while ready_to_refresh:  
            try:
                Workbook.RefreshAll()
                ready_to_refresh = False
            except (AttributeError,com_error):
                time.sleep(1)
                print('Workbook Refresh Error')
                ready_to_refresh = True
    
        ready_to_sync = True
        while ready_to_sync:
            try:
                xlApp.CalculateUntilAsyncQueriesDone()
                ready_to_sync = False
            except (AttributeError, com_error):
                time.sleep(1)
                print('Workbook Sync Error')
                ready_to_sync = True

        Process_Data()

        ready_to_save = True 
        while ready_to_save:
            try:
                Workbook.Save()
                Workbook.Close()
                Workbook = None
                ready_to_save = False
            except (com_error, AttributeError):  
                time.sleep(1)
                print('Workbook Save Error')
                ready_to_save = True
        xlApp.Quit()
### please change path(The path where you'll need to set up your path for your excel, and better, much storing it )
path = "C:\\Users\\JakeBondoa\\Desktop\\Python excel\\Dataexcelup.xlsx"
while True:
    Refresh_Save()
