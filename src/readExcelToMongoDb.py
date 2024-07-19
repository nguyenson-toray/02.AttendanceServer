import pandas as pd
from pymongo import MongoClient
import schedule
import time
from datetime import datetime
import sys
import numpy as np

# Connect to MongoDB
connection_string = "mongodb://localhost:27017/"
database_name = "tiqn"
client = MongoClient(connection_string)
db = client[database_name]
collectionEmployee = db["Employee"]
# Excel file path (replace with your actual path)
excel_file = r"\\fs\tiqn\03.Department\01.Operation Management\03.HR-GA\01.HR\Toray's employees information All in one.xlsx"
# excel_file = r"D:\Programming\01.AttendanceApp\02.Server\Toray's employees information All in one.xlsx"
def excelAllInOneToMongoDb() -> int:
    # Read data from Excel using pandas
    data = pd.read_excel(excel_file, keep_default_na=False, na_values='', na_filter=False)
    data.fillna("", inplace=True)
    # Convert pandas dataframe to dictionary (adjust based on your data structure)
    data_dict = data.to_dict(orient="records")
    # Loop through each data row
    count = 0
    for row in data_dict:
        count += 1
        # Update document in MongoDB collection based on a unique identifier (replace with your logic)
        if row["Emp Code"] == '' or row["Emp Code"] == 0 or row["Fullname"] == '' or row["Fullname"] == 0 or row['_id'] == 0 or row['_id'] == '':
            print(f'BYPASS ROW - Empty Emp Code or Fullname or _id: {row}')
            continue
        if row["Working/Resigned"] == 0 or row["Working/Resigned"] == '':
            print(f'BYPASS ROW - Resigned: {row}')
            continue
        if row["Fullname"] == 'Shoji Izumi':
            print(f'BYPASS ROW - Shoji Izumi: {row}')
            continue
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
        print(f"collection.update_one: {update}")
        collectionEmployee.update_one(filter, update, upsert=True)  # Upsert inserts if no document matches the filter
    return count

if __name__ == "__main__":
    print(f"Data updated in MongoDB collection : {excelAllInOneToMongoDb()} records")
    # Close the connection
    client.close()
