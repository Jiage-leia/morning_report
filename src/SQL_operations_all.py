# a simple database interface for Python that builds on top of freeTDs to provide a Python DB-API interface to Microsoft SQL server.
import pymssql
# import MySQLdb
# from MySQLdb import _mysql
#the SQL AST vocabulary allows SQL code abstract syntax trees to be published in RDF (resource description framework)
import ast
#json stands for Javascript Object Notation, it is mainly used in storing and transporting data
#yaml helps to translate data format 
import json, yaml
import re

# set time format
from datetime import datetime as dt, timedelta

# connect python with sql
def generate_connector_of_MS_SQL(ip, user, password, db_name):
    return pymssql.connect(ip, user, password, db_name, charset="utf8")

# define "IOV" appointment
def find_daily_iov(conn, schema, date=None):
    #define query date
    if not date:
        _tomorrow = dt.now().date() + timedelta(1)
    else:
        _tomorrow = date
    # print(_tomorrow)
    #define sql
    _sql = "select reason, status, CONVERT(time, appt_time) as appt_time " \
           "from {0}.appointment_view " \
           "where appt_date = '{1}' and reason like 'IOV%' " \
           "order by appt_time ;" \
           ";".format(schema, _tomorrow)

    try:
        cursor = conn.cursor()
        cursor.execute(_sql)
        return cursor.fetchall()
    except:
        return ()

# define "talk" appoinment
def find_daily_talk(conn, schema, date=None):
    #define query date
    if not date:
        _tomorrow = dt.now().date() + timedelta(1)
    else:
        _tomorrow = date
    #define sql
    _sql_talk = "select reason, status, CONVERT(time, appt_time) as appt_time " \
           "from {0}.appointment_view " \
           "where appt_date = '{1}' and (reason like 'PHONE%' " \
           "or reason like 'TELE%' or reason like 'CON%' " \
           "or reason like 'COUR%' or reason like '%TALK') " \
           "order by appt_time ;" \
           ";".format(schema, _tomorrow)

# cursor is an object which helps to execute the query and fetch the records from the database
    try:
        cursor = conn.cursor()
        cursor.execute(_sql_talk)
        return cursor.fetchall()
    except:
        return ()

# define "nppa" appoinment
def find_daily_nppa(conn, schema, date=None):
    #define query date
    if not date:
        _tomorrow = dt.now().date() + timedelta(1)
    else:
        _tomorrow = date
    #define sql
    _sql_nppa = "select reason, status, CONVERT(time, appt_time) as appt_time " \
           "from {0}.appointment_view " \
           "where appt_date = '{1}' and ( reason like 'IUI%' " \
           "or reason like '%HSG%' "\
           "or reason like 'R1%' "\
           "or reason like 'SIS%' "\
           "order by appt_time;" \
           ";".format(schema, _tomorrow)

# cursor is an object which helps to execute the query and fetch the records from the database
    try:
        cursor = conn.cursor()
        cursor.execute(_sql_nppa)
        return cursor.fetchall()
    except:
        return ()