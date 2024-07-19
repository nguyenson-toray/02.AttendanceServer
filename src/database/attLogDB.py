from src.database.attLogMachine import AttLogMachine


class AttLogDB(AttLogMachine):
    def __init__(self,machineNo, uid, user_id, name, empId):
        super.__init__(self,machineNo, uid, user_id, name)
        self.empId=empId
    def __str__(self):
        return f"Machine: {self.machineNo}  uid: {self.uid} user_id: {self.user_id}  name: {self.name}  empId: {self.empId}"