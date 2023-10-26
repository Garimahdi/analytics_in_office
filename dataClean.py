# To Clean Data for Pelindo Reporting

import pandas as pd
import datetime
from openpyxl import load_workbook

def dataClean(file):
    file.drop(columns=[
    'black_toner',
    'cyan_toner',
    'magenta_toner',
    'yellow_toner',
    'black_drum',
    'cyan_drum',
    'magenta_drum',
    'yellow_drum',
    'waste_toner',
    'sn_black_toner',
    'sn_cyan_toner',
    'sn_magenta_toner',
    'sn_yellow_toner',
    'sn_black_drum',
    'sn_cyan_drum',
    'sn_magenta_drum',
    'sn_yellow_drum'
    ], inplace=True)
    
    file['log_date'] = pd.to_datetime(file['log_date'])
    file['log_date'] = file['log_date'].apply(lambda x: x.strftime('%Y-%m-%d'))
    return file

def gapSearch(data):
    bef_bp = data['black_printed'][len(data['black_printed'])-1]
    bef_bc = data['black_copied'][len(data['black_copied'])-1]
    bef_cp = data['color_printed'][len(data['color_printed'])-1]
    bef_cc = data['color_copied'][len(data['color_copied'])-1]
    
    aft_bp = data['black_printed'][0]
    aft_bc = data['black_copied'][0]
    aft_cp = data['color_printed'][0]
    aft_cc = data['color_copied'][0]
    
    bw = (aft_bp + aft_bc)-(bef_bp + bef_bc)
    clr = (aft_cp + aft_cc)-(bef_cp + bef_cc)
    
    return bw, clr

def dataReader(txt):
    df = pd.read_csv(r'{}'.format(txt))
    df['model_code'] = df['serial_number'].apply(lambda x:'JL'+x[:8])
    df['sn'] = df['serial_number'].apply(lambda x:x[-6:])
    df.drop(columns='serial_number', inplace=True)
    return df

def xlsxWriter(*data, path, sheetName):
    x = path
    book = load_workbook(x)
    writer = pd.ExcelWriter(x, engine='openpyxl')
    writer.book = book
    data.to_excel(writer, sheet_name=sheetName)
    writer.closed()