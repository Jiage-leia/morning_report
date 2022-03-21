Default_Cycle_Length = 28

LPS_Within_Days = 15

LPS_1st_IVF_Within_Days = 5

Cycle_Count_Dict_Date_limit_Num = 2

# Max_Trig_ER_Interval = 5

Error_Code_Number = -1

Error_Code_Cycle_Length = Error_Code_Number

Error_Code_Age = 0

Error_Code_Cycle_Date = 0

Error_Code_String = ""

Error_Code_Date = "NULL"

Identical_Delta_Number = 1

Uniform_Med_Dose_Level_H = 1000

Egg_Retrieval_Delta_Days_Max = 4

E2_Increase_A_Lot = 1.8

Abs_Start_Time = "2001-01-01 00:00:00"

# ====================================================================================================

# Project_Root_DIR = "/home/aiadmin/emr_ai_full_func/ai_4_fertility_center/ai_4_fertility_center/db_update"
Project_Root_DIR = "db_update"

db_update_Dir_Last_Folder = "db_update"

# Path_STD = "/home/aiadmin/emr_ai_full_func/ai_4_fertility_center/ai_4_fertility_center/db_update/log/"
Path_STD = "/Users/eva/Documents/research/IVF/emr_ai6/ai_4_fertility_center/ai_4_fertility_center/db_update/log/"

# Path_EMB = '/media/WindowsShared/Embryo_Inventory.xlsx'
Path_EMB = '/Users/eva/Documents/research/IVF/emr_ai6/EMBRYO_RECORDS/EMB_INV/Embryo_Inventory.xlsx'

# Path_EMB_HIST = '/media/WindowsShared/Embryo_Inventory_Hist.xlsx'
Path_EMB_HIST = '/Users/eva/Documents/research/IVF/emr_ai6/EMBRYO_RECORDS/EMB_INV/Embryo_Inventory_Hist.xlsx'

Path_NICS = '/Users/eva/Documents/research/IVF/emr_ai6/EMBRYO_RECORDS/NICS/NICS_RECORDS.xlsx'
# Path_NICS = '/media/WindowsShared/NICS.xlsx'
NICS_Table_Name = 'Samples'
NICS_Sheet_Name = 'Samples'
BIOPSY_Table_Name = 'LOGC'

Path_PGT = '/Users/eva/Documents/research/IVF/emr_ai6/EMBRYO_RECORDS/PGT/'
# Path_PGT = '/media/WindowsShared/PGT/'
PGT_Sheet_Name = 'LOG ENTRY'

EMB_INV_LOG_PATH = '/Users/eva/Documents/research/IVF/emr_ai6/ai_4_fertility_center/ai_4_fertility_center/db_update/log/EMB_INV/'
# EMB_INV_LOG_PATH = '/home/aiadmin/emr_ai_full_func/ai_4_fertility_center/ai_4_fertility_center/db_update/log/EMB_INV/'

Is_MSSQL = True

if Is_MSSQL:
    IP_Server = "192.168.193.251\\NHFC_ART_CLIENT"
else:
    IP_Server = "192.168.1.112"

if Is_MSSQL:
    User_Server = "aireader"
else:
    User_Server = "root"

#Oct/16/20 Jia: ART_NHFC readonly account
ART_NHFC_User_readonly = "jwang"
ART_NHFC_Password_readonly = "xAr2wew7!"
#

if Is_MSSQL:
    PWD_Server = "Int3llig3nCe!"
else:
    PWD_Server = "root"

if Is_MSSQL:
    DB_Name_Server = "ART_NHFC"
else:
    DB_Name_Server = "dbo"

Schema_Server = "dbo"

IP_Local = "localhost"

User_Local = "root"

# PWD_Local = "3561768!Zn"
# PWD_Local = "root"
PWD_Local = "NewHopeAIVF"

Schema_IVF = "nhfc"

Schema_Wrt = Schema_IVF

Static_Work = True
Static_Mode = 'full'
Static_Partial_Mode_Date_Range = 7

Online_Work = not Static_Work

Batch_Work = False

# Static_Work_End_Time = "2019-09-30 23:59:59"
Static_Work_End_Time = "2020-03-25 09:21:00"

CycleCount_Based = True

Schedule_Time = "01:00:00"

Schedule_Work = "schedule_work"

RealTime_Work = "realtime_work"

APPOINTMENT_CODE_MAP = {
  0: 'Pending',
  1: 'Confirmed',
  2: 'Waiting',
  3: 'BeingSeen',
  4: 'Completed',
  5: 'Late',
  6: 'Missed',
  7: 'Canceled',
  8: 'Rescheduled',
  9: 'Recalled',
  10: 'Unknown',
};


if __name__ == "__main__":
    print('global_variables.py')


