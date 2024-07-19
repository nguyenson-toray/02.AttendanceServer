from multiprocessing import Process
import pymongo
import sys
import os
from datetime import datetime

sys.path.insert(1, os.path.abspath("./pyzk"))
from zk import ZK, const

client = pymongo.MongoClient("mongodb://localhost:27017/")
db = client["tiqn"]  # Replace "mydatabase" with your database name
collectionEmployee = db["Employee"]
collectionAttLogMachine = db["AttLogMachine"]
collectionAttLog = db["AttLog"]
collectionHistoryGetAttLogs = db["HistoryGetAttLogs"]
# Find all documents
allEmployee = collectionEmployee.find()
historyGetAttLogs = collectionHistoryGetAttLogs.find()


def get_att_log_one_time(machine: ZK, machineNo: int) -> int:
    conn = None
    count = 0
    timeBeginGetLogs = datetime.now()
    try:

        history = {}
        for his in collectionHistoryGetAttLogs.find():
            if (int(his['machine']) == machineNo):
                history = his
        lastTime = history['lastTimeGetAttLogs']
        conn = machine.connect()
        print(f"Connecting machine : {machine.get_network_params()['ip']}")
        # disable device, this method ensures no activity on the device while the process is run
        conn.disable_device()
        # another commands will be here!
        # Get attendances (will return list of Attendance object)
        attendances = conn.get_attendance()
        for attendance in attendances:
            if (attendance.timestamp > lastTime):
                mydict = {"machineNo": machineNo, "uid": attendance.uid, "attFingerId": int(attendance.user_id),
                          "empId": 'No Emp Id', "name": 'No name',
                          "timestamp": attendance.timestamp}
                for emp in collectionEmployee.find():
                    if (int(emp['attFingerId']) == int(attendance.user_id)):
                        mydict = {"machineNo": machineNo, "uid": attendance.uid, "attFingerId": int(attendance.user_id),
                                  "empId": emp['empId'], "name": emp['name'],
                                  "timestamp": attendance.timestamp}
                        break
                collectionAttLog.insert_one(mydict)
                count += 1
        myquery = {"machine": machineNo}
        newvalue = {"$set": {"lastTimeGetAttLogs": datetime.now(), "lastCount": count}}
        collectionHistoryGetAttLogs.update_one(myquery, newvalue)
        conn.enable_device()
    except:
        print('except')
    finally:
        if conn:
            conn.disconnect()
        print(f"Machine {machineNo} => {count} records. Total time: {datetime.now() - timeBeginGetLogs}")

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
        Process(target=get_att_log_one_time, args=(machine, machineNo)).start()
