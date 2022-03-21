# import MySQLdb
# # Just for Mac OS debugging env setup.
# import pymssql
import pymysql
pymysql.install_as_MySQLdb()
import MySQLdb

# set time
from datetime import datetime as dt, timedelta
# from xlwt import Workbook
from openpyxl import Workbook, load_workbook

#import definitions of variables
from global_variables import *
#import SQL definition
from SQL_operations_all import *

import json
import os
import re

if __name__ == "__main__":
    # HRS_PATTERN = r'.*\s+after\s+(\S+)\s+Hours'
    # FILE_NAME = '/media/WindowsShared/list_pre_ER_patient_for_LAB_DPS.xlsx'
    FILE_NAME = './test.xlsx'
    
    
    #connect mssql database
    conn_dbo = generate_connector_of_MS_SQL(IP_Server, User_Server, PWD_Server, DB_Name_Server)
    #get date obj of tomorrow
    tomorrow_date = dt.now().date() + timedelta(1)
    #query data IOV
    dataIOV = find_daily_iov(conn_dbo, Schema_Server)
    #query data Talk
    dataTALK = find_daily_talk(conn_dbo, Schema_Server)
    #query data nppa
    dataNPPA = find_daily_nppa(conn_dbo, Schema_Server)

    #call excel api
    if os.path.exists(FILE_NAME):
        wb = load_workbook(filename=FILE_NAME, read_only=False)
        ws = wb.active
        ws.delete_rows(1, amount=ws.max_row)

    else:
        wb = Workbook()
        ws = wb.active
        ws.title = 'DPS'

#define excel header
    ws.append(['Chart ID', 'Date', 'E2', 'PR-2', '# of Follicles: L side', '# of Follicles: R side', 'Trigger Method', 'Trigger Time', 'ER After Trigger'])

#data preprocessing for IOV
    _iov_dict = {'ZHANG': [], 'LIU': [], 'WONG': [], 'ZEITOUN': [], 'MAKAROV': []}
    for row in dataIOV:
        _platform = None
        _reason, _status, _time = row
        _status = APPOINTMENT_CODE_MAP[_status]

        if _reason.find('PH') != -1:
            _platform = "Phone"
        elif _reason.find('SK') != -1:
            _platform = "Skype"
        else:
            _platform = "In Person"

        if _reason.find("ZHAN") != -1:
            _iov_dict['ZHANG'].append([_time, _status, _platform])
            continue

        if _reason.find("LIU") != -1:
            _iov_dict['LIU'].append([_time, _status, _platform])
            continue

        if _reason.find("WONG") != -1:
            _iov_dict['WONG'].append([_time, _status, _platform])
            continue

        if _reason.find("ZEIT") != -1:
            _iov_dict['ZEITOUN'].append([_time, _status, _platform])
            continue

        if _reason.find("MAKA") != -1:
            _iov_dict['MAKAROV'].append([_time, _status, _platform])
            continue


# data processing for talk
    _talk_dict = {'ZHANG': [], 'LIU': [], 'WONG': [], 'ZEITOUN': [], 'MAKAROV': []}
    for row in dataTALK:
        _platform = None
        _reason, _status, _time = row
        _status = APPOINTMENT_CODE_MAP[_status]

        if _reason.find('PH') != -1:
            _platform = "Phone"
        elif _reason.find('SCAN') != -1:
            _platform = "Scan"
        elif _reason.find('TELE') != -1:
            _platform = "TeleHealth"
        elif _reason.find('CON') != -1:
            _platform = "Consultation"
        elif _reason.find('ENDO') != -1:
            _platform = "Endosee"
        elif _reason.find('COUR') != -1:
            _platform = "Courtesy Talk"
        

        if _reason.find("ZHAN") != -1:
            _talk_dict['ZHANG'].append([_time, _status, _platform])
            continue

        if _reason.find("LIU") != -1:
            _talk_dict['LIU'].append([_time, _status, _platform])
            continue

        if _reason.find("WONG") != -1:
            _talk_dict['WONG'].append([_time, _status, _platform])
            continue

        if _reason.find("ZEIT") != -1:
            _talk_dict['ZEITOUN'].append([_time, _status, _platform])
            continue

        if _reason.find("MAKA") != -1:
            _talk_dict['MAKAROV'].append([_time, _status, _platform])
            continue


#data processing for Nurse Practitioner/Physician assistant 
    _nppa_dict = {'IUI': [], 'HSG': [], 'R1': [], 'SIS': []}
    for row in dataNPPA:
        _reason, _status, _time = row
        _status = APPOINTMENT_CODE_MAP[_status]


        if _reason.find("IUI") != -1:
            _nppa_dict['IUI'].append([_time, _status])
            continue

        if _reason.find("HSG") != -1:
            _nppa_dict['HSG'].append([_time, _status])
            continue

        if _reason.find("R1") != -1:
            _nppa_dict['R1'].append([_time, _status])
            continue

        if _reason.find("SIS") != -1:
            _nppa_dict['SIS'].append([_time, _status])
            continue

    # row_counter = 1
    # for d in data:
    #     chart_id, date_, e2, p4, raw_follicle_data, trigger_method = d
    #     trig, trig_time, hrs = '', '', ''
    #     follicle_data = json.loads(raw_follicle_data)
    #     follicle_g14_counter_l = 0
    #     follicle_g14_counter_r = 0
    #     for key in follicle_data:
    #         if follicle_data[key] != '':
    #             sub_f_l = follicle_data[key].split(', ')
    #             for f in sub_f_l:
    #                 if float(f) >= 14:
    #                     if key == "L":
    #                         follicle_g14_counter_l += 1
    #                     else:
    #                         follicle_g14_counter_r += 1
        
    #     if trigger_method.find('after') != -1:
    #         m = re.match(HRS_PATTERN, trigger_method)
    #         hrs = int(m.group(1))
        
    #     if trigger_method.find('Lupron') != -1 and trigger_method.find('Ovidrel') != -1:
    #         trig = 'BOT'
        
    #     if trigger_method.find('Lupron') != -1 and trigger_method.find('Ovidrel') == -1:
    #         trig = 'LUP'

    #     if trigger_method.find('Lupron') == -1 and trigger_method.find('Ovidrel') != -1:
    #         trig = 'OVI'

    #     if trigger_method.find('Trigger Time') != -1:
    #         trig_time = trigger_method[trigger_method.find('Trigger Time')+12 : trigger_method.find(' and ER')]

    #     if follicle_g14_counter_l > 0 or follicle_g14_counter_r > 0:
    #         ws.append([chart_id, date_.strftime('%Y-%m-%d'), e2, p4, follicle_g14_counter_l, follicle_g14_counter_r, trig, trig_time, hrs])
    #         row_counter += 1
    
    wb.save(FILE_NAME) 
    conn_dbo.close()
