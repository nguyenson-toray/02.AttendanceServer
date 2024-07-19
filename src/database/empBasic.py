class EmpBasic:
    def __init__(self,id, attFingerId, department, directIndirect, dob, empId, gender, group, joiningDate, level, lineTeam, maternityComebackDate, name, positionE, resignDate, section, sewingNonSewing, shift, supporting, workStatus ):
        self.id = id
        self.attFingerId = attFingerId
        self.department = department
        self.directIndirect = directIndirect
        self.dob = dob
        self.empId = empId
        self.name = name
        self.gender = gender
        self.group = group
        self.joiningDate = joiningDate
        self.level = level
        self.lineTeam = lineTeam
        self.maternityComebackDate = maternityComebackDate
        self.positionE = positionE
        self.resignDate = resignDate
        self.section = section
        self.sewingNonSewing = sewingNonSewing
        self.shift = shift
        self.supporting = supporting
        self.workStatus = workStatus
    def __str__(self):
        return f"id: {self.id} attFingerId: {self.attFingerId} empId: {self.empId} name: {self.name}"
