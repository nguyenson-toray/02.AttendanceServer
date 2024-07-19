# -*- coding: utf-8 -*-
import os
import sys
import time
from multiprocessing import Process
import os
import schedule
from zk import ZK
import pymongo
from datetime import datetime
from datetime import timedelta
import pandas as pd

CWD = os.path.dirname(os.path.realpath(__file__))
ROOT_DIR = os.path.dirname(CWD)
sys.path.append(ROOT_DIR)

client = pymongo.MongoClient("mongodb://localhost:27017/")
db = client["tiqn"]  # Replace "mydatabase" with your database name
collectionEmployee = db["Employee"]
collectionAttLog = db["AttLog"]
collectionMaternityTracking = db["MaternityTracking"]
collectionHistoryGetAttLogs = db["HistoryGetAttLogs"]
collectionLastModifiedExcelData = db["LastModifiedExcelData"]

def floor(datetimeObj):
   return datetimeObj.replace()
def excelAllInOneToMongoDb()-> bool:
    # Read data from Excel using pandas
    excelFile = r"\\fs\tiqn\03.Department\01.Operation Management\03.HR-GA\01.HR\Toray's employees information All in one.xlsx"
    lastModifiedFile = datetime.fromtimestamp(os.path.getmtime(excelFile)).replace(microsecond=0)
    lastModifiedDb = collectionLastModifiedExcelData.find_one()['aio'].replace(microsecond=0)
    needUpdate = True if lastModifiedFile > lastModifiedDb else False
    if not needUpdate:
        print(f"{excelFile} => No change from {lastModifiedDb} => pass")
        return needUpdate
    print(f"{excelFile} => Changed at {lastModifiedFile} => Need update")
    query = {'_id': 1}
    newValue = {"$set": {"aio": lastModifiedFile}}
    collectionLastModifiedExcelData.update_one(query, newValue)
    data = pd.read_excel(excelFile, keep_default_na=False, na_values='', na_filter=False)
    data.fillna("", inplace=True)
    # Convert pandas dataframe to dictionary (adjust based on your data structure)
    data_dict = data.to_dict(orient="records")
    # Loop through each data row
    count = 0
    for row in data_dict:

        # Update document in MongoDB collection based on a unique identifier (replace with your logic)
        if row["Emp Code"] == '' or row["Emp Code"] == 0 or row["Fullname"] == '' or row["Fullname"] == 0 or row[
            '_id'] == 0 or row['_id'] == '':
            # print(f'BYPASS ROW - Empty Emp Code or Fullname or _id: {row}')
            continue
        if row["Fullname"] == 'Shoji Izumi':
            # print(f'BYPASS ROW - Shoji Izumi: {row}')
            continue
        count += 1
        filter = {"_id": row["_id"]}  # Assuming "_id" is a unique identifier in your data
        # update = {"$set": row}  # Update all fields in the document
        fieldCollected = {}
        empId = 'TIQN-XXXX' if row['Emp Code'] == '' else row['Emp Code']
        name = 'No name' if row['Fullname'] == '' else row['Fullname']
        attFingerId = 0 if row['Finger Id'] == '' else row['Finger Id']
        department = '' if row['Department'] == '' else row['Department']
        section = '' if row['Section'] == '' else row['Section']
        group = '' if row['Group'] == '' else row['Group']
        lineTeam = '' if row['Line/ Team'] == '' else row['Line/ Team']
        gender = '' if row['Gender'] == '' else row['Gender']
        position = '' if row['Position'] == '' else row['Position']
        level = '' if row['Level'] == '' else row['Level']
        directIndirect = '' if row['Direct/ Indirect'] == '' else row['Direct/ Indirect']
        sewingNonSewing = '' if row['Sewing/Non sewing'] == '' else row['Sewing/Non sewing']
        supporting = '' if row['Supporting'] == '' else row['Supporting']
        dob = datetime.fromisoformat('1900-01-01') if str(type(row['DOB'])).find('datetime.datetime') == -1 else row[
            'DOB']
        joiningDate = datetime.fromisoformat('1900-01-01') if str(type(row['Joining date'])).find(
            'datetime.datetime') == -1 else row['Joining date']
        workStatus = 'Resigned' if (row['Working/Resigned'] == 0 or row['Working/Resigned'] == '') else 'Working'
        resignOn = datetime.fromisoformat('2099-01-01')
        fieldCollected.update(
            {'empId': empId, 'name': name, 'attFingerId': attFingerId, 'department': department, 'section': section,
             'group': group, 'lineTeam': lineTeam, 'gender': gender, 'position': position, 'level': level,
             'directIndirect': directIndirect, 'sewingNonSewing': sewingNonSewing, 'supporting': supporting, 'dob': dob,
             'joiningDate': joiningDate, 'workStatus': workStatus, 'resignOn': resignOn})
        update = {"$set": fieldCollected}
        # print(f"collection.update_one: {update}")
        collectionEmployee.update_one(filter, update, upsert=True)  # Upsert inserts if no document matches the filter
    print(f" excelAllInOneToMongoDb - {datetime.now()} : {count} records")
    return needUpdate


def excelMaternityToMongoDb(forceUpdate: bool):
    excelFile = r"\\fs\tiqn\03.Department\01.Operation Management\03.HR-GA\05.Nurse\Maternity leave\Danh sách nhân viên nữ mang thai. 1.xlsx"
    # check if file changed
    lastModifiedFile = datetime.fromtimestamp(os.path.getmtime(excelFile)).replace(microsecond=0)
    lastModifiedDb = collectionLastModifiedExcelData.find_one()['maternity'].replace(microsecond=0)
    needUpdateMaternity = forceUpdate if forceUpdate else True if lastModifiedFile > lastModifiedDb else False
    if not needUpdateMaternity:
        print(f"{excelFile} =>No need update or no change from {lastModifiedDb} => pass")
        return
    print(f"{excelFile} =>Need to update, changed at {lastModifiedFile} => Need update")
    query = {'_id': 1}
    newValue = {"$set": {"maternity": lastModifiedFile}}
    collectionLastModifiedExcelData.update_one(query, newValue)
    # Read -thai san
    data = pd.read_excel(excelFile, sheet_name='Thai sản', keep_default_na=False, na_values='',
                         na_filter=False, skiprows=3)
    data.fillna("", inplace=True)
    # Convert pandas dataframe to dictionary (adjust based on your data structure)
    data_dict = data.to_dict(orient="records")
    # Loop through each data row
    empIdMaternity = 'TIQN-XXXX'
    for row in data_dict:
        if row['STT'] != 0 and row['STT'] != '':
            empIdMaternity = row['MSNV']
            query = {'empId': empIdMaternity}
            update = {"$set": {'workStatus': 'Maternity leave'}}
            print(f"update {empIdMaternity} to Maternity leave")
            collectionEmployee.update_one(query, update)

    ##---------mang thai--pregnant-------
    data = pd.read_excel(excelFile, sheet_name='mang thai', keep_default_na=False, na_values='',
                         na_filter=False, skiprows=3)
    data.fillna("", inplace=True)
    # Convert pandas dataframe to dictionary (adjust based on your data structure)
    data_dict = data.to_dict(orient="records")
    # Loop through each data row
    for row in data_dict:
        if row['STT'] != 0 and row['STT'] != '':
            empIdMaternity = row['MSNV']
            query = {'empId': empIdMaternity}
            update = {"$set": {'workStatus': 'Working pregnant'}}
            collectionEmployee.update_one(query, update)
            print(f"update {empIdMaternity} to Working pregnant")
    ##---------Con nhỏ dưới 12 tháng---------
    data = pd.read_excel(excelFile, sheet_name='Con nhỏ dưới 12 tháng', keep_default_na=False, na_values='',
                         na_filter=False, skiprows=3)
    data.fillna("", inplace=True)
    # Convert pandas dataframe to dictionary (adjust based on your data structure)
    data_dict = data.to_dict(orient="records")
    # Loop through each data row
    for row in data_dict:
        if row['STT'] != 0 and row['STT'] != '':
            empIdMaternity = row['MSNV']
            query = {'empId': empIdMaternity}
            update = {"$set": {'workStatus': 'Working young child'}}
            print(f"update {empIdMaternity} to Working young child")
            collectionEmployee.update_one(query, update)


def excelResignToMongoDb(forceUpdate: bool):
    excelFile = r"\\fs\tiqn\03.Department\01.Operation Management\03.HR-GA\01.HR\7 Resigned list\Resigned report.xlsx"
    # check if file changed
    lastModifiedFile = datetime.fromtimestamp(os.path.getmtime(excelFile)).replace(microsecond=0)
    lastModifiedDb = collectionLastModifiedExcelData.find_one()['resign'].replace(microsecond=0)
    needUpdateResign =forceUpdate if forceUpdate else True if lastModifiedFile > lastModifiedDb else False
    if not needUpdateResign:
        print(f"{excelFile} =>No need update or No change from {lastModifiedDb} => pass")
        return forceUpdate
    print(f"{excelFile} =>Need to update, Changed at {lastModifiedFile} => Need update")
    query = {'_id': 1}
    newValue = {"$set": {"resign": lastModifiedFile}}
    collectionLastModifiedExcelData.update_one(query, newValue)
    data = pd.read_excel(excelFile, sheet_name='QD', keep_default_na=False, na_values='',
                         na_filter=False)
    data.fillna("", inplace=True)
    # Convert pandas dataframe to dictionary (adjust based on your data structure)
    data_dict = data.to_dict(orient="records")
    # Loop through each data row
    for row in data_dict:
        if row['Số QĐ'] != 0 and row['Số QĐ'] != '':
            empIdResigned = row['MSNV']
            resignDate = datetime.fromisoformat('2099-01-01') if (str(type(row['Ngày nghỉ việc'])).find(
                'datetime.datetime') == -1 and str(type(row['Ngày nghỉ việc'])).find(
                'timestamps.Timestamp') == -1) else row['Ngày nghỉ việc']
            query = {'empId': empIdResigned}
            update = {"$set": {'workStatus': 'Resigned', 'resignOn': resignDate}}
            print(f"update {empIdResigned} to resigned on {resignDate}")
            collectionEmployee.update_one(query, update)
    return forceUpdate
def live_capture_attendance(machine: ZK, machineNo) -> None:
    conn = None
    try:
        conn = machine.connect()
        print(f"Live capture attendance: {machine.get_network_params()['ip']}")
        for attendance in conn.live_capture():
            if attendance is None:
                pass
            else:
                mydict = {"machineNo": machineNo, "uid": attendance.uid, "attFingerId": int(attendance.user_id),
                          "empId": 'No Emp Id', "name": 'No name',
                          "timestamp": attendance.timestamp}
                for emp in collectionEmployee.find():
                    if (int(emp['attFingerId']) == int(attendance.user_id)):
                        mydict = {"machineNo": machineNo, "uid": attendance.uid,
                                  "attFingerId": int(attendance.user_id),
                                  "empId": emp['empId'], "name": emp['name'],
                                  "timestamp": attendance.timestamp}
                        break
                print(mydict)
                collectionAttLog.insert_one(mydict)
                myquery = {"machine": machineNo}
                newvalue = {"$set": {"lastTimeGetAttLogs": datetime.now(), "lastCount": 1}}
                collectionHistoryGetAttLogs.update_one(myquery, newvalue)

    except Exception as e:
        print("Process terminate : {}".format(e))
    finally:
        if conn:
            conn.disconnect()

def updateExceltoMongoDb():
    needUpdate= excelAllInOneToMongoDb()
    excelMaternityToMongoDb(needUpdate)
    excelResignToMongoDb(needUpdate)

if __name__ == "__main__":  # confirms that the code is under main function

    machine1 = ZK('192.168.1.31', port=4370, timeout=5, password=0, force_udp=False, ommit_ping=False)
    machine2 = ZK('192.168.1.32', port=4370, timeout=5, password=0, force_udp=False, ommit_ping=False)
    machine3 = ZK('192.168.1.33', port=4370, timeout=5, password=0, force_udp=False, ommit_ping=False)
    machine4 = ZK('192.168.1.34', port=4370, timeout=5, password=0, force_udp=False, ommit_ping=False)
    attMachines = [machine1, machine2, machine3, machine4]

    # Read excel
    updateExceltoMongoDb()
    # Read Att log, excel
    machineNo = 0
    for machine in attMachines:
        machineNo += 1
        Process(target=live_capture_attendance, args=(machine, machineNo)).start()
        # Create a new process
    schedule.every().minute.do(updateExceltoMongoDb)
    while True:
        schedule.run_pending()
        time.sleep(1)
