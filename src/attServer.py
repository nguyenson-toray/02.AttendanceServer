# -*- coding: utf-8 -*-
import sys
import time
import threading
import os
import schedule
from zk import ZK
import pymongo
from datetime import datetime, timedelta
import pandas as pd
import cv2
from pyzbar.pyzbar import decode
from pdf2image import convert_from_path
from openpyxl import load_workbook

CWD = os.path.dirname(os.path.realpath(__file__))
ROOT_DIR = os.path.dirname(CWD)
sys.path.append(ROOT_DIR)

client = pymongo.MongoClient("mongodb://localhost:27017/")
db = client["tiqn"]
collection_employee = db["Employee"]
collection_att_log = db["AttLog"]
collection_maternity_tracking = db["MaternityTracking"]
collection_history_get_att_logs = db["HistoryGetAttLogs"]
collection_last_modified_data_hr = db["LastModifiedDataHR"]
collection_ot_register = db["OtRegister"]
list_emp = []
ip_fs = r'\\192.168.1.13'
ip_att_machines = []
path_config = {}


def write_log(log):
    log_str = str(log)
    try:
        file_name = r"..\03.Logs\log_" + datetime.now().strftime("%Y%m%d") + ".txt"
        with open(file_name, 'a', encoding="utf-8") as file:
            file.writelines(log_str)
            file.close()
    except Exception as e:
        print(f'!!!!!!!!! write_log Exception :{e}') if enable_print else None



def excel_aio_to_db() -> bool:
    # Read data from Excel using pandas
    excel_file = path_config['aio']
    log = ''
    last_modified_file = datetime.fromtimestamp(os.path.getmtime(excel_file)).replace(microsecond=0)
    last_modified_db = collection_last_modified_data_hr.find_one()['aio'].replace(microsecond=0)
    need_update = True if last_modified_file > last_modified_db else False
    if not need_update:
        # print(f"    No change from {last_modified_db} => Pass") if enable_print else None
        return need_update
    log += f"{datetime.now()} excel_aio_to_db\n    Changed at {last_modified_file} => Need update\n"
    print(f"{datetime.now()} excel_aio_to_db\n    Changed at {last_modified_file} => Need update") if enable_print else None
    query = {'_id': 1}
    new_value = {"$set": {"aio": last_modified_file}}
    collection_last_modified_data_hr.update_one(query, new_value)
    data = pd.read_excel(excel_file, keep_default_na=False, na_values='', na_filter=False)
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
        if row["Fullname"] == 'Shoji Izumi' or row["Fullname"] == 'Amagata Osamu':
            count += 1
            log += f'   BYPASS Izumi, Amagata\n'
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
        log += f'   update #{count}: {empId}     {name}\n'
        # print(f"collection.update_one: {update}")
        collection_employee.update_one(filter, update, upsert=True)  # Upsert inserts if no document matches the filter
    log += f"    {count} records\n"
    print(f"    {count} records") if enable_print else None
    write_log(log)
    return need_update


def excel_maternity_to_db(forceUpdate: bool):
    log = ''
    excel_file = path_config['maternity']
    # check if file changed
    last_modified_file = datetime.fromtimestamp(os.path.getmtime(excel_file)).replace(microsecond=0)
    last_modified_db = collection_last_modified_data_hr.find_one()['maternity'].replace(microsecond=0)
    needUpdateMaternity = forceUpdate if forceUpdate else True if last_modified_file > last_modified_db else False
    if not needUpdateMaternity:
        return
    log += f"{datetime.now()} excelMaternityToMongoDb\n   Need to update, changed at {last_modified_file} => Need update\n"
    print(
        f"{datetime.now()} excelMaternityToMongoDb\n    Need to update, changed at {last_modified_file} => Need update") if enable_print else None
    query = {'_id': 1}
    new_value = {"$set": {"maternity": last_modified_file}}
    collection_last_modified_data_hr.update_one(query, new_value)
    # Read -thai san
    data = pd.read_excel(excel_file, sheet_name='Thai sản', keep_default_na=False, na_values='',
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
            maternityLeaveBegin = datetime.fromisoformat('2099-01-01') if (str(type(row['NGÀY NGHỈ SINH'])).find(
                'datetime.datetime') == -1 and str(type(row['NGÀY NGHỈ SINH'])).find(
                'timestamps.Timestamp') == -1) else row['NGÀY NGHỈ SINH']
            maternityLeaveEnd = datetime.fromisoformat('2099-01-01') if (str(type(row['NGÀY QUAY LẠI'])).find(
                'datetime.datetime') == -1 and str(type(row['NGÀY QUAY LẠI'])).find(
                'timestamps.Timestamp') == -1) else row['NGÀY QUAY LẠI']
            maternityLeaveEnd = maternityLeaveEnd - timedelta(days=1)
            if (maternityLeaveBegin.year != 2099 and maternityLeaveEnd.year != 2099):
                update = {"$set": {'workStatus': 'Maternity leave', 'maternityLeaveBegin': maternityLeaveBegin,
                                   'maternityLeaveEnd': maternityLeaveEnd}}
            else:
                update = {"$set": {'workStatus': 'Maternity leave'}}
            log += f'   update {empIdMaternity} to Maternity leave\n'
            print(f"    update {empIdMaternity} to Maternity leave") if enable_print else None
            collection_employee.update_one(query, update)

    ##---------mang thai--pregnant-------
    data = pd.read_excel(excel_file, sheet_name='mang thai', keep_default_na=False, na_values='',
                         na_filter=False, skiprows=3)
    data.fillna("", inplace=True)
    # Convert pandas dataframe to dictionary (adjust based on your data structure)
    data_dict = data.to_dict(orient="records")
    # Loop through each data row
    for row in data_dict:
        if row['STT'] != 0 and row['STT'] != '':
            empIdMaternity = row['MSNV']
            maternityBegin = row['NGÀY NHẬN THÔNG TIN']
            maternityEnd = row['NGÀY DỰ SINH']
            query = {'empId': empIdMaternity}
            update = {"$set": {'workStatus': 'Working pregnant', 'maternityBegin': maternityBegin,
                               'maternityEnd': maternityEnd}}
            collection_employee.update_one(query, update)
            log += f"   update {empIdMaternity} to Working pregnant\n"
            print(f"    update {empIdMaternity} to Working pregnant") if enable_print else None
    ##---------Con nhỏ dưới 12 tháng---------
    data = pd.read_excel(excel_file, sheet_name='Con nhỏ dưới 12 tháng', keep_default_na=False, na_values='',
                         na_filter=False, skiprows=3)
    data.fillna("", inplace=True)
    # Convert pandas dataframe to dictionary (adjust based on your data structure)
    data_dict = data.to_dict(orient="records")
    # Loop through each data row
    for row in data_dict:
        if row['STT'] != 0 and row['STT'] != '':
            empIdMaternity = row['MSNV']
            maternityBegin = row['NGÀY QUAY LẠI']
            maternityEnd = row['NGÀY CUỐI CÙNG THỜI GIAN NUÔI CON NHỎ']
            query = {'empId': empIdMaternity}
            update = {"$set": {'workStatus': 'Working young child', 'maternityBegin': maternityBegin,
                               'maternityEnd': maternityEnd}}
            log += f"   update {empIdMaternity} to Working young child\n"
            print(f"   update {empIdMaternity} to Working young child") if enable_print else None
            collection_employee.update_one(query, update)
    write_log(log)


def excel_resign_to_db(forceUpdate: bool):
    log = ''
    excel_file = path_config['resign']
    last_modified_file = datetime.fromtimestamp(os.path.getmtime(excel_file)).replace(microsecond=0)
    last_modified_db = collection_last_modified_data_hr.find_one()['resign'].replace(microsecond=0)
    needUpdateResign = forceUpdate if forceUpdate else True if last_modified_file > last_modified_db else False
    if not needUpdateResign:
        return forceUpdate
    log += f"{datetime.now()} excelMaternityToMongoDb\n    Need to update, Changed at {last_modified_file} => Need update\n"
    print(
        f"{datetime.now()} excelMaternityToMongoDb\n    Need to update, Changed at {last_modified_file} => Need update") if enable_print else None
    query = {'_id': 1}
    new_value = {"$set": {"resign": last_modified_file}}
    collection_last_modified_data_hr.update_one(query, new_value)
    data = pd.read_excel(excel_file, sheet_name='QD', keep_default_na=False, na_values='',
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

            if resignDate.year > 1900 and resignDate < datetime.now():
                query = {'empId': empIdResigned}
                update = {"$set": {'workStatus': 'Resigned', 'resignOn': resignDate}}
                log += f"    update {empIdResigned} to resigned on {resignDate}\n"
                print(f"    update {empIdResigned} to resigned on {resignDate}") if enable_print else None
                collection_employee.update_one(query, update)
    write_log(log)
    return forceUpdate


def get_att_log_one_time(machine: ZK, machineNo: int) -> int:
    conn = None
    count = 0
    log = f'{datetime.now()} get_att_log_one_time: machine {machineNo}\n'
    print(f'{datetime.now()} get_att_log_one_time: machine {machineNo}')
    time_begin_get_logs = datetime.now()
    try:

        history = {}
        for his in collection_history_get_att_logs.find():
            if int(his['machine']) == machineNo:
                history = his
        last_time_db = history['lastTimeGetAttLogs']
        last_time_machine=last_time_db
        conn = machine.connect()
        print(f"Connecting machine : {machineNo} ...")
        # disable device, this method ensures no activity on the device while the process is run
        conn.disable_device()
        # another commands will be here!
        # Get attendances (will return list of Attendance object)
        attendances = conn.get_attendance()
        for attendance in attendances:
            if attendance.timestamp > last_time_db:
                mydict = {"machineNo": machineNo, "uid": attendance.uid, "attFingerId": int(attendance.user_id),
                          "empId": 'No Emp Id', "name": 'No name',
                          "timestamp": attendance.timestamp}
                emp_id = find_emp_id_by_finger_id(int(attendance.user_id))
                emp_name = find_name_by_finger_id(int(attendance.user_id))
                if emp_id != 'Not found':
                    mydict['empId'] = emp_id
                if emp_name != 'Not found':
                    mydict['name'] = emp_name
                log1 = f'    Add: Machine: {mydict['machineNo']}       {mydict['timestamp'].strftime("%d-%m-%Y %H:%M:%S")}      {mydict['attFingerId']}     {mydict['empId']}    {mydict['name']}\n'
                collection_att_log.insert_one(mydict)
                print(log1)
                log += log1
                count += 1
                last_time_machine = attendance.timestamp
        my_query = {"machine": machineNo}
        new_value = {"$set": {"lastTimeGetAttLogs": last_time_machine, "lastCount": count}}
        collection_history_get_att_logs.update_one(my_query, new_value)
        conn.enable_device()
    except Exception as e:
        print(f'     exception: {e}')
        write_log(f'     exception: {e}')
    finally:
        if conn:
            conn.disconnect()
        log += f'    Machine {machineNo} => {count} records. Total time: {datetime.now() - time_begin_get_logs}\n'
        print(
            f"    Machine {machineNo} => {count} records. Total time: {datetime.now() - time_begin_get_logs}")
    write_log(log)


def live_capture_attendance(machine: ZK, machineNo) -> None:
    conn = None
    log = f'{datetime.now()} live_capture_attendance: machine {machineNo}\n'
    write_log(log)
    print(f'{datetime.now()} live_capture_attendance: machine {machineNo}') if enable_print else None
    try:
        conn = machine.connect()
        for attendance in conn.live_capture():
            if attendance is None:
                pass
            else:
                mydict = {"machineNo": machineNo, "uid": attendance.uid, "attFingerId": int(attendance.user_id),
                          "empId": 'No Emp Id', "name": 'No name',
                          "timestamp": attendance.timestamp}

                emp_id = find_emp_id_by_finger_id(int(attendance.user_id))
                emp_name = find_name_by_finger_id(int(attendance.user_id))
                if emp_id != 'Not found':
                    mydict['empId'] = emp_id
                if emp_name != 'Not found':
                    mydict['name'] = emp_name
                log = f'    live_capture add: Machine: {mydict['machineNo']}      {mydict['timestamp'].strftime("%d-%m-%Y %H:%M:%S")}     {mydict['attFingerId']}    {mydict['empId']}   {mydict['name']}\n'
                collection_att_log.insert_one(mydict)
                print(log)
                write_log(log)
                myquery = {"machine": machineNo}
                new_value = {"$set": {"lastTimeGetAttLogs": attendance.timestamp, "lastCount": 1}}
                collection_history_get_att_logs.update_one(myquery, new_value)

    except Exception as e:
        print("     Exception: Process terminate : {}".format(e))
        write_log(f'    live_capture Exception: {e}')
    finally:
        if conn:
            conn.disconnect()


def update_excel_to_mongoDb():
    current_time = datetime.now()
    read_config()
    if (current_time.hour == 17 and current_time.minute < 30) or (current_time.hour == 16 and current_time.minute > 55):
        log = f'{current_time} is PEAK TIME => bypass update_excel_to_mongoDb'
        print(log) if enable_print else None
        write_log(log)
        return
    need_update = excel_aio_to_db()
    excel_maternity_to_db(need_update)
    excel_resign_to_db(need_update)
    if need_update:
        get_list_emp()


def get_list_emp():
    list_emp.clear()
    count = 0
    for emp in collection_employee.find():
        count += 1
        list_emp.append({'attFingerId': emp['attFingerId'], 'empId': emp['empId'], 'name': emp['name']})
    print(f'{datetime.now()} get_list_emp: Total:{count}\n') if enable_print else None
    log = f'{datetime.now()} get_list_emp:\n{list_emp} \nTotal:{count}\n'

    excel_file = r"..\02.Config\config.xlsx"
    print(f'{datetime.now()} read_config : {excel_file}') if enable_print else None
    data_phu_quy = pd.read_excel(excel_file, sheet_name='phú quý')
    data_dict_phu_quy = data_phu_quy.to_dict(orient="records")
    for row in data_dict_phu_quy:
        list_emp.append({'attFingerId': row['FingerID'], 'empId': row['Code'], 'name': row['Name']})
    write_log(log)


def ot_register_detect_qr_and_save():
    log = ''
    begin = datetime.now()
    ot_folder_path = path_config['ot_folder']
    ot_folder_pdf_imported_path = ot_folder_path + r"\01.Imported"
    ot_folder_pdf_path = ot_folder_path + r"\02.Pdf"
    ot_excel_file_name = "01.OT summary.xlsx"
    my_poppler_path = r'C:\Program Files\poppler-24.02.0\Library\bin'
    ot_request_no_on_db = []
    for doc in collection_ot_register.find():
        ot_request_no_on_db.append(str(doc['requestNo']))

    ot_last_request_id_db = 0
    ot_last_ot_register_document = collection_ot_register.find().sort({'_id': -1}).limit(1)
    for doc in ot_last_ot_register_document:
        ot_last_request_id_db = doc['_id']
    ot_all_file_names_registed = []
    for doc in collection_last_modified_data_hr.find():
        ot_all_file_names_registed.append(str(doc['otRequestFileName']))
    ot_all_file_names_registed_str = str(ot_all_file_names_registed[0])
    # print(f'ot_all_file_names_registed_str : {ot_all_file_names_registed_str}') if enable_print else None
    pdf_files = []
    try:
        # Use glob with wildcard for efficient PDF search
        for filename in os.listdir(ot_folder_pdf_path):
            if filename.endswith('.pdf'):  # and filename.startswith("202"):
                pdf_files.append(filename)
                # pdf_paths.append(os.path.join(ot_folder_pdf_path, filename))
        # Sort PDFs by last modified time (descending) using reverse=True
        # pdf_files.sort(key=os.path.getmtime)
        # print(f"pdf_files: {pdf_files}")
    except Exception as e:
        log += f'   Exception : Error processing folder {ot_folder_pdf_path}: {e}\n'
        print(f"    Exception : Error processing folder {ot_folder_pdf_path}: {e}")
    for file_pdf in pdf_files:
        # file_last_modified = datetime.fromtimestamp(os.path.getmtime(file_path))
        print(f'*** Check file_pdf : {file_pdf}') if enable_print else None
        # print(f'File: {file_path} => file_last_modified : {file_last_modified}')
        if (not ot_all_file_names_registed_str.__contains__(file_pdf)):
            file_path = os.path.join(ot_folder_pdf_path, file_pdf)
            print(f'    => NOT YET ADDED  :{file_path}') if enable_print else None
            log += f'    Read QR code in file: {file_path}\n'
            try:
                temp_pages = convert_from_path(file_path, dpi=300, output_file='qr.png', paths_only=True,
                                               output_folder=ot_folder_pdf_path, poppler_path=my_poppler_path)
                for temp_page in temp_pages:
                    try:
                        page = cv2.imread(temp_page)
                        gray = cv2.cvtColor(page, cv2.COLOR_BGR2GRAY)  # Convert to grayscale for better detection
                        # Detect QR codes
                        qr_codes = decode(gray)
                        # Extract data from QR codes
                        for qr_code in qr_codes:
                            qr = str(qr_code.data.decode('utf-8'))
                            print(f'File : {file_pdf} :    Detected QR code: {qr}\n') if enable_print else None
                            log += f'File : {file_pdf} :    Detected QR code: {qr}\n'
                            ot_request_no = qr.split(';')[0].strip()
                            date_request = qr.split(';')[1].strip()
                            list_date_time = qr.split(';')[2].strip().split(', ')
                            list_emp_id = qr.split(';')[3].strip().split(' ')
                            # ot_last_request_time = file_last_modified
                            if (not ot_request_no_on_db.__contains__(ot_request_no)):
                                ot_request_no_on_db.append(ot_request_no)
                                ot_last_request_id_db = qr_code_ot_register_to_db(ot_last_request_id_db, ot_request_no,
                                                                                  date_request,
                                                                                  list_date_time, list_emp_id)
                    except Exception as e:
                        log += f'   detect_qr : Error 1 :{e}\n'
                        print(f"detect_qr : Error 1 :{e}") if enable_print else None

            except Exception as e:
                log += f'   detect_qr : Error 2 :{e}\n'
                print(f"detect_qr : Error 2 :{e}") if enable_print else None
            query = {'_id': 1}
            ot_all_file_names_registed_str = ot_all_file_names_registed_str + '; ' + file_pdf
            update = {
                "$set": {'otRequestFileName': ot_all_file_names_registed_str,
                         # "otRequestId": ot_last_request_id_db
                         }}
            collection_last_modified_data_hr.update_one(query, update)
            src_path = os.path.join(ot_folder_pdf_path, file_pdf)
            dst_path = os.path.join(ot_folder_pdf_imported_path, file_pdf)
            os.rename(src_path, dst_path)  # move file pdf scaned to "01.Imported"
            log += f'Scanned & moved to folder "01.Imported" : {file_pdf}'
            print(f'   => add {file_pdf} to DB : ') if enable_print else None
            for filename in os.listdir(ot_folder_pdf_path):
                file_path = os.path.join(ot_folder_pdf_path, filename)
                if filename.endswith('.ppm'):
                    os.remove(file_path)
                    print(f'    delete .ppm : {filename}') if enable_print else None

        else:
            print(f'  => Pass') if enable_print else None
    try:
        ot_register_append_excel(ot_folder_path, ot_excel_file_name)
    except Exception as e:
        log += f'   detect_qr Exception: Error 3 :{e}\n'
        print(f"    detect_qr Exception: Error 3 :{e}")

    # end = datetime.now()
    # time = end - begin
    # print(f'    Total time scan QR : {time.total_seconds()}') if enable_print else None
    # log += f'   Total time scan QR : {time.total_seconds()}\n'
    write_log(log)


def ot_register_append_excel(folder_path: str, file_name: str):
    filename = os.path.join(folder_path, file_name)
    ot_last_request_id_db = 0
    ot_last_ot_register_document = collection_ot_register.find().sort({'_id': -1}).limit(1)
    for doc in ot_last_ot_register_document:
        ot_last_request_id_db = doc['_id']
    if ot_last_request_id_db == 0:
        return
    # Open the existing workbook
    wb = load_workbook(filename=filename)
    sheet_name = "data"  # Replace with your sheet name
    # Select the sheet you want to append data to (adjust sheet name)
    ws = wb[sheet_name]

    if ws.max_row == 1:
        ot_last_request_id_excel = 0
        last_row = 1
    else:
        last_row = ws.max_row
        ot_last_request_id_excel = int(ws.cell(last_row, 1).value)
    if ot_last_request_id_db > ot_last_request_id_excel:
        log = f'{datetime.now()} ot_register_append_excel\n'
        print(F'{datetime.now()} Append OT data from db to excel') if enable_print else None
        query = {"_id": {"$gt": ot_last_request_id_excel}}
        docs = collection_ot_register.find(query)
        for doc in docs:
            log += f"   Append ID : {doc['_id']}\n"
            print(f"   Append ID : {doc['_id']}\n") if enable_print else None
            last_row += 1
            ws.cell(row=last_row, column=1).value = doc['_id']
            ws.cell(row=last_row, column=2).value = doc['requestNo']
            ws.cell(row=last_row, column=3).value = doc['requestDate']
            ws.cell(row=last_row, column=4).value = doc['otDate']
            ws.cell(row=last_row, column=5).value = doc['otTimeBegin']
            ws.cell(row=last_row, column=6).value = doc['otTimeEnd']
            ws.cell(row=last_row, column=7).value = doc['empId']
            ws.cell(row=last_row, column=8).value = doc['name']
        # Apply date formatting to the entire column
        date_format = "DD-MM-YYYY"
        for row in ws.iter_rows(min_row=2):  # Skip the first row (headers)
            cell = row[2]  # Access cell by column index (0-based)
            if isinstance(cell.value, datetime):  # Check if value is a date object
                cell.number_format = date_format
            else:
                pass
            cell = row[3]  # Access cell by column index (0-based)
            if isinstance(cell.value, datetime):  # Check if value is a date object
                cell.number_format = date_format
            else:
                pass
        # Save the modified workbook
        wb.save(filename=filename)
        wb.close()
        print(f"  => ot_register_append_excel : Data appended to: {filename}") if enable_print else None
        write_log(log)


def qr_code_ot_register_to_db(ot_last_request_id_db: int, ot_request_no: str, date_request: str, list_date_time: list,
                              list_emp_id: list) -> int:
    log = f'{datetime.now()} qr_code_ot_register_to_db: {ot_request_no}     List Date & Time: {list_date_time} List Emp ID: {list_emp_id}\n'
    print(
        f'{datetime.now()} qr_code_ot_register_to_db: {ot_request_no}     List Date & Time: {list_date_time} List Emp ID: {list_emp_id}\n') if enable_print else None
    for date_and_time in list_date_time:
        if not len(date_and_time) == 20 or date_and_time.find('19000100') >= 0:
            print(f'{date_and_time}: is wrong => pass') if enable_print else None
            continue
        date_ot_str = date_and_time.strip().split(' ')[0]
        ot_time_begin_str = date_and_time.strip().split(' ')[1]
        ot_time_end_str = date_and_time.strip().split(' ')[2]
        ot_date = datetime.strptime(date_ot_str[0:8], '%Y%m%d')
        request_date = datetime.strptime(date_request, '%Y%m%d')
        for id_no in list_emp_id:
            emp_id = 'TIQN-' + str(id_no)
            emp_name = find_name_by_emp_id(emp_id)
            ot_last_request_id_db += 1
            mydict = {"_id": ot_last_request_id_db, "requestNo": ot_request_no, "requestDate": request_date,
                      "otDate": ot_date, 'otTimeBegin': ot_time_begin_str, 'otTimeEnd': ot_time_end_str,
                      "empId": emp_id, "name": emp_name}
            log += f'   Add record OT register: No: {ot_request_no}      Request date: {mydict['requestDate']}      Date: {mydict['otDate']}       From: {mydict['otTimeBegin']}       To: {mydict['otTimeEnd']}       Emp ID: {mydict['empId']}       Name: {mydict['name']}\n'
            collection_ot_register.insert_one(mydict)
    write_log(log)
    return ot_last_request_id_db


def find_name_by_emp_id(emp_id: str) -> str:
    name = 'Not found'
    for emp in list_emp:
        if emp['empId'] == emp_id:
            name = emp['name']
    return name


def find_name_by_finger_id(finger_id: int) -> str:
    name = 'Not found'
    for emp in list_emp:
        if emp['attFingerId'] == finger_id:
            name = emp['name']
            break
    return name


def find_emp_id_by_finger_id(finger_id: int) -> str:
    emp_id = 'Not found'
    for emp in list_emp:
        if emp['attFingerId'] == finger_id:
            emp_id = emp['empId']
            break
    return emp_id


def get_att_log(machine: ZK, machineNo, real_time: bool):
    get_att_log_one_time(machine, machineNo)
    if real_time:
        live_capture_attendance(machine, machineNo)


def sync_time_devices():
    log = f'{datetime.now()} sync_time_devices\n'
    try:
        for ip in ip_att_machines:
            machine = ZK(ip, port=4370, timeout=5, password=0, force_udp=False, ommit_ping=False)
            conn = machine.connect()
            time = datetime.now()
            print(f"    Machine : {ip}\n    Current time : {conn.get_time()}    Sync time to : {time}")
            conn.set_time(time)
    except Exception as e:
        log += f'   Sync time ERROR: {e}\n'
        print("     Sync time ERROR : {}".format(e))
    write_log(log)


def read_config():
    ip_att_machines.clear()
    log = f'{datetime.now()} read_config:\n'
    excel_file = r"..\02.Config\config.xlsx"
    print(f'{datetime.now()} read_config : {excel_file}') if enable_print else None
    data_ip = pd.read_excel(excel_file, sheet_name='ip_machines')
    data_dict_ip = data_ip.to_dict(orient="records")
    for row in data_dict_ip:
        ip_att_machines.append(row['IP'])
    log = f'     ip_att_machines: {ip_att_machines}'
    print(f'    ip_att_machines: {ip_att_machines}') if enable_print else None
    data_path = pd.read_excel(excel_file, sheet_name='path')
    data_dict_path = data_path.to_dict(orient="records")
    path_config.clear()
    for row in data_dict_path:
        if row['Name'] == 'aio':
            path_config['aio'] = row['Path']
        if row['Name'] == 'resign':
            path_config['resign'] = row['Path']
        if row['Name'] == 'maternity':
            path_config['maternity'] = row['Path']
        if row['Name'] == 'ot_folder':
            path_config['ot_folder'] = row['Path']
        if row['Name'] == 'maternity_pregnant':
            path_config['maternity_pregnant'] = row['Path']
        if row['Name'] == 'maternity_young_child':
            path_config['maternity_young_child'] = row['Path']
        if row['Name'] == 'maternity_leave':
            path_config['maternity_leave'] = row['Path']
    print(f'    path_config: {path_config}') if enable_print else None
    log += f'      path_config: {str(path_config)}'
    write_log(log)


if __name__ == "__main__":
    main_log = f'---------------------------*** ATTENDANCE SERVER ***-----------------------------------------\n'
    main_log += f'START AT : {datetime.now().strftime("%d-%m-%Y %H:%M:%S")}\n'
    main_log += ''' 
    schedule.every().sunday.at("06:00").do(sync_time_devices)
    schedule.every().day.at("09:00").do(update_excel_to_mongoDb)
    schedule.every().day.at("11:55").do(update_excel_to_mongoDb)
    schedule.every().day.at("15:00").do(update_excel_to_mongoDb)
    schedule.every().day.at("18:00").do(update_excel_to_mongoDb)
    schedule.every(10).minutes.do(ot_register_detect_qr_and_save)'''
    main_log += f'\n-------------------------------------------------------------------------------------------\n'
    print(main_log)
    write_log(main_log)
    enable_print = False
    update_excel_to_mongoDb()
    get_list_emp()
    machine_no = 1
    for ip in ip_att_machines:
        machine = ZK(ip, port=4370, timeout=5, password=0, force_udp=False, ommit_ping=False)
        threading.Thread(target=get_att_log, args=(
            machine, machine_no, True), name=ip).start()
        machine_no += 1
    ot_register_detect_qr_and_save()
    # -----------------------
    schedule.every().sunday.at("06:00").do(sync_time_devices)
    schedule.every().day.at("09:00").do(update_excel_to_mongoDb)
    schedule.every().day.at("11:55").do(update_excel_to_mongoDb)
    schedule.every().day.at("15:00").do(update_excel_to_mongoDb)
    schedule.every().day.at("18:00").do(update_excel_to_mongoDb)
    schedule.every(10).minutes.do(ot_register_detect_qr_and_save)
    while True:
        schedule.run_pending()
        time.sleep(1)
