class AttLogMachine:
    def __init__(self,machineNo, uid, user_id, timestamp):
        self.machineNo = machineNo
        self.uid= uid
        self.attFingerId = user_id
        self.timestamp = timestamp
    def __str__(self):
        return f"Machine: {self.machineNo}  uid: {self.uid} attFingerId: {self.attFingerId} timestamp:{self.timestamp}"


