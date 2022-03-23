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
#import excel helper
from processing_excel import *

import json
import os
import re

if __name__ == "__main__":
    # HRS_PATTERN = r'.*\s+after\s+(\S+)\s+Hours'
    # FILE_NAME = '/media/WindowsShared/list_pre_ER_patient_for_LAB_DPS.xlsx'
    BASIC_PATH = '/Users/eva/Documents/research/IVF/emr_ai6/ai_4_fertility_center/ai_4_fertility_center/morning_report/'
    TEMPLATE_FILE_NAME = BASIC_PATH+'Morning Schedule.xlsx'
    
    
    
    #connect mssql database
    conn_dbo = generate_connector_of_MS_SQL(IP_Server, User_Server, PWD_Server, DB_Name_Server)
    #get date obj of tomorrow
    tomorrow_date = dt.now().date() + timedelta(1)

    OUT_FILE_NAME = BASIC_PATH + 'Morning Schedule' + tomorrow_date.strftime("%Y-%m-%d") + '.xlsx'
    #query data IOV
    dataIOV = find_daily_iov(conn_dbo, Schema_Server)
    #query data Talk
    dataTALK = find_daily_talk(conn_dbo, Schema_Server)
    #query data nppa
    dataNPPA = find_daily_nppa(conn_dbo, Schema_Server)

    #call excel api
    assert os.path.exists(TEMPLATE_FILE_NAME)

    wb = load_workbook(filename=TEMPLATE_FILE_NAME, read_only=False)
    ws = wb.active

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

        if _reason.find("MAKA") != -1 or _reason.find("JKM") != -1:
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
        # elif _reason.find('SCAN') != -1:
        #     _platform = "Scan"
        elif _reason.find('TELE') != -1:
            _platform = "TeleHealth"
        elif _reason.find('CON') != -1:
            _platform = "Consultation"
        # elif _reason.find('ENDO') != -1:
        #     _platform = "Endosee"
        elif _reason.find('COUR') != -1:
            _platform = "Courtesy Talk"
        elif _reason.find('TALK') != -1:
            _platform = "Talk"
        

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

        if _reason.find("MAKA") != -1 or _reason.find("JKM") != -1:
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

    conn_dbo.close()
    for md in MDs:
        ws = insert_md_row(ws, md, _iov_dict[md], _talk_dict[md])

    _nppa_appt_dict = {name: [] for name in PA_NP}
    
    # not used until pa np code started to be applied.
    # for pa_np in PA_NP:
    #     ws = insert_pa_np_row(ws, pa_np, _nppa_appt_dict)

    ws = insert_task_row(ws, _nppa_appt_dict, _nppa_dict)
    wb.save(filename=OUT_FILE_NAME)
