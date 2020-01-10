import DWDataReader
import datetime
import pandas as pd
import numpy as np
from pandas.io.excel import ExcelWriter
import matplotlib.pyplot as plt


#dll = DWDataReader.open_dll()
#%% Functions

def downSampleDataFrame(dataFrame, ms):
    # only use for downsampling, upsampling will result in N/A values
    sampleTime= str(ms) + 'ms'
    # Resampling methods sum, mean, std, sem, max, min, median, first, last, ohlc:
    resampled = dataFrame.resample(sampleTime).mean()
    print("Resampled with sample time of", sampleTime, "\n")
    return resampled

def volt2stress(measuredVolt, k=2, v0=5, eModulusMPa=210000):
    #Consumes Volt, returns, stress in MPa
    rek = ( 4 / k ) * ( v0 / measuredVolt )
    stress = rek * eModulusMPa
    return stress

def deweFileInfo(deweFile):
    fields = []
    df = DWDataReader.read_dws(deweFile, fields, rename, dll=dll)
    if isinstance(df,dict):
        for k, v in df.items():
            if not isinstance(v,list):
                # [print(i) for i in v]
                print(k,' ',v)
            else:
                lists = v
                for i in lists:
                    print("Channel ",i[0], "; Name:", i[1], "; Renamed:", i[2], "; Unit:", i[3], "; Sample count:", i[4], "; Sample rate factor:", i[5] , sep="")

def exportXlsx(dataFrame):
    #Write Excel file
    excelWriteFile = "dewe2excel_" + "_" + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + ".xlsx"
    with ExcelWriter(excelWriteFile) as xlsxWriter:
        dataFrame.between_time(intervalStart, intervalStop).to_excel(xlsxWriter)

#%%

deweFile = 'FrieseSluis_v6_2019_10_31_213000.dxz'

fields = []
rename = {  b'Rek_01':'Rek_01',
            b'Rek_02':'Rek_02',
            b'Rek_03':'Rek_03',
            b'Rek_04':'Rek_04',
            b'Rek_05':'Rek_05',
            b'Rek_06':'Rek_06',
            b'Rek_07':'Rek_07',
            b'Rek_08':'Rek_08',
            b'Rek_09':'Rek_09',
            b'Rek_10':'Rek_10',
            b'Rek_11':'Rek_11',
            b'Rek_12':'Rek_12',
            b'CPU (AVE)' :'CPU (AVE)',
            b'MemFree':'MemFree',
            b'DiskFree':'DiskFree' }

deweFile = 'FrieseSluis_v5_2019_10_31_173000.dxz'

dll = DWDataReader.open_dll()
df = DWDataReader.read_dws(deweFile, fields, rename, dll=dll)

''' Reads fields from a DEWESoft file and returns a pandas DataFrame
    
    Args:
        - filename: input filename
        
        - fields: A list of fields to extract. If not specified, information
            about the file will be returned. Empty list means read all fields.
            Note that selection is performed before renaming, so names refer
            to original field names from the DEWESoft file.
            
        - rename: A dict used to rename fields in the resulting DataFrame. If
            the key is an integer, it is assumed that this is the zero-based
            index of the field to rename. Otherwise it is assumed to be the
            original field name.
        
        - mixed_sample_rate: If false, only fields with identical sample rates
            will be allowed. The sampling rate of the returned data will be the
            sample rate of the selected fields. E.g., if the file contains
            fields sampled at 100Hz, 20Hz and 10Hz, and you select fields with
            20Hz, this will be the sampling rate of the returned data.
            
            If true, mixed sample rates will be allowed and the missing values
            will be filled with NaN-s. The sample rate in this case will be the
            maximum of all fields (e.g., 100Hz in the previous example).
            
            Note that only integer sample rate ratios are handled.
            
        - dll: optional handle to dll, obtained with open_dll(). Use when
            calling this function many times, to prevent opening/closing the
            DLL.
'''            
DWDataReader.close_dll(dll)

'''  Returns:
        If there is no 'fields' input, returns a dict with file info:
            - sampling_rate (integer)
            - start_store_time (datetime)
            - duration (seconds)
            - number_of_channels (int)
            - channel_info: list of tuples, for each channel:
                - channel index (int), zero-based
                - channel name (str)
                - renamed channel name (str), possibly the same as channel name
                - unit (str)
                - number of samples (int)
                - sampling rate ratio divisor (int), 1 for full sampling rate
'''
# If fields is empty print info

if isinstance(df,pd.DataFrame):   
    print("\n")
    # Drop columns if required
    #dropColumns = {}
    dropColumns = [df.columns[2], df.columns[3], df.columns[8], df.columns[11]]
    # for i in dropColumns:
        # df.drop(columns=[i])
    df = df.drop(columns=dropColumns)
    df = df.drop(columns='Rek_07')

    # print(df.describe(include='all'))

    df = downSampleDataFrame(df, 20)

    #Substract mean of the series to remove offset the signal
    for col in df.columns:
        mean = df[col].mean()
        # print (col, mean)
        df[col] = df[col].sub(mean)
        # print (col, dataFrame[col].mean())

    print(df.describe(include='all'))
    

    intervalStart = '17:33:02'
    intervalStop  = '17:45:00'
    plt.close('all')
    df.between_time(intervalStart, intervalStop).plot()    
    plt.show()
    
exportXlsx(df.between_time(intervalStart, intervalStop))

