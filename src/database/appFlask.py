from flask import Flask, jsonify
from flask_pymongo import PyMongo
from datetime import datetime
app = Flask(__name__)
app.config["MONGO_URI"] = "mongodb://localhost:27017/tiqn"
mongo = PyMongo(app)
@app.route("/employees")
def read_collection_employee():
    print('0--------------')
    # Get the collection
    collection = mongo.db.Employee  # Replace with your collection name
    # Read all documents (adjust query for filtering)
    results = collection.find()
    data=[]
    for element in results:
        element.pop("_id")
        dob=element['dob'].strftime('%Y-%m-%d')
        joiningDate=element['joiningDate'].strftime('%Y-%m-%d')
        maternityComebackDate=element['maternityComebackDate'].strftime('%Y-%m-%d')
        resignDate=element['resignDate'].strftime('%Y-%m-%d')
        element.pop("dob")
        element.pop("joiningDate")
        element.pop("maternityComebackDate")
        element.pop("resignDate")
        element['dob']=dob
        element['joiningDate']=joiningDate
        element['maternityComebackDate']=maternityComebackDate
        element['resignDate']=resignDate
        data.append(element)
    # Return data as JSON
    print(f"read_collection_employee  => result {len(data)} records")
    return jsonify(data)
@app.route("/attLogs/<beginTimeStr>/<endTimeStr>")
def read_collection_att_log(beginTimeStr,endTimeStr):
    collection = mongo.db.AttLog
    beginTime= datetime(int(beginTimeStr[0:4]), int(beginTimeStr[4:6]), int(beginTimeStr[6:8]), 0, 0,0)
    endTime=datetime(int(endTimeStr[0:4]), int(endTimeStr[4:6]), int(endTimeStr[6:8]), 23, 59,59)
    results = collection.find({'timestamp': {'$gte': beginTime, '$lte': endTime}})
    data=[]
    for element in results:
        element.pop("_id")
        data.append(element)
    print(f"read_collection_att_log : From: {beginTime} to {endTime} => result {len(data)} records")
    return jsonify(data)
if __name__ == '__main__':
    app.run()