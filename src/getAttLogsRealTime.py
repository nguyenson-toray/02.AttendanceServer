# -*- coding: utf-8 -*-
import os
import sys
import time
from multiprocessing import Process

import schedule
from zk import ZK
import pymongo
from datetime import datetime
import pandas as pd
CWD = os.path.dirname(os.path.realpath(__file__))
ROOT_DIR = os.path.dirname(CWD)
sys.path.append(ROOT_DIR)

client = pymongo.MongoClient("mongodb://localhost:27017/")
db = client["tiqn"]  # Replace "mydatabase" with your database name
collectionEmployee = db["Employee"]
collectionAttLog = db["AttLog"]
collectionHistoryGetAttLogs = db["HistoryGetAttLogs"]
excel_file = r"\\fs\tiqn\03.Department\01.Operation Management\03.HR-GA\01.HR\Toray's employees information All in one.xlsx"
def excelAllInOneToMongoDb():
    # Read data from Excel using pandas
    data = pd.read_excel(excel_file, keep_default_na=False, na_values='', na_filter=False)
    data.fillna("", inplace=True)
    # Convert pandas dataframe to dictionary (adjust based on your data structure)
    data_dict = data.to_dict(orient="records")
    # Loop through each data row
    count = 0
    for row in data_dict:

        # Update document in MongoDB collection based on a unique identifier (replace with your logic)
        if row["Emp Code"] == '' or row["Emp Code"] == 0 or row["Fullname"] == '' or row["Fullname"] == 0 or row['_id'] == 0 or row['_id'] == '':
            # print(f'BYPASS ROW - Empty Emp Code or Fullname or _id: {row}')
            continue
        if row["Working/Resigned"] == 0 or row["Working/Resigned"] == '':
            # print(f'BYPASS ROW - Resigned: {row}')
            continue
        if row["Fullname"] == 'Shoji Izumi':
            # print(f'BYPASS ROW - Shoji Izumi: {row}')
            continue
        count += 1
        filter = {"_id": row["_id"]}  # Assuming "_id" is a unique identifier in your data
        update = {"$set": row}  # Update all fields in the document
        fieldCollected = {}
        fieldCollected.update({'empId': 'TIQN-XXXX' if row['Emp Code'] == '' else row['Emp Code']})
        fieldCollected.update({'name': 'No Name' if row['Fullname'] == '' else row['Fullname']})
        fieldCollected.update({'attFingerId': 0 if row['Finger Id'] == '' else row['Finger Id']})
        fieldCollected.update({'department': '' if row['Department'] == '' else row['Department']})
        fieldCollected.update({'section': '' if row['Section'] == '' else row['Section']})
        fieldCollected.update({'group': '' if row['Group'] == '' else row['Group']})
        fieldCollected.update({'lineTeam': '' if row['Line/ Team'] == '' else row['Line/ Team']})
        fieldCollected.update({'gender': 'F' if row['Gender'] == '' else row['Gender']})
        fieldCollected.update({'position': '' if row['Position'] == '' else row['Position']})
        fieldCollected.update({'level': '' if row['Level'] == '' else row['Level']})
        fieldCollected.update({'directIndirect': '' if row['Direct/ Indirect'] == '' else row['Direct/ Indirect']})
        fieldCollected.update({'sewingNonSewing': '' if row['Sewing/Non sewing'] == '' else row['Sewing/Non sewing']})
        fieldCollected.update({'supporting': '' if row['Supporting'] == '' else row['Supporting']})
        fieldCollected.update({'dob': row['DOB']})
        fieldCollected.update({'joiningDate': row['Joining date']})
        fieldCollected.update({'workStatus': 0 if row['Working/Resigned'] == '' else row['Working/Resigned']})
        fieldCollected.update({'maternity': 0 if row['Maternity'] == '' else row['Maternity']})
        update = {"$set": fieldCollected}
        # print(f"collection.update_one: {update}")
        collectionEmployee.update_one(filter, update, upsert=True)  # Upsert inserts if no document matches the filter
    print(f" excelAllInOneToMongoDb - {datetime.now()} : {count} records")
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


if __name__ == "__main__":  # confirms that the code is under main function

    machine1 = ZK('192.168.1.31', port=4370, timeout=5, password=0, force_udp=False, ommit_ping=False)
    machine2 = ZK('192.168.1.32', port=4370, timeout=5, password=0, force_udp=False, ommit_ping=False)
    machine3 = ZK('192.168.1.33', port=4370, timeout=5, password=0, force_udp=False, ommit_ping=False)
    machine4 = ZK('192.168.1.34', port=4370, timeout=5, password=0, force_udp=False, ommit_ping=False)
    attMachines = [machine1, machine2, machine3, machine4]
    # Create a new process
    # for machine in attMachines:
    machineNo = 0
    for machine in attMachines:
        machineNo += 1
        Process(target=live_capture_attendance, args=(machine, machineNo)).start()

    schedule.every().minute.do(excelAllInOneToMongoDb)
    while True:
        schedule.run_pending()
        time.sleep(1)
