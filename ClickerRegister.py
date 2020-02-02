from zipfile import ZipFile as _ZipFile 
from xml.dom import minidom as _minidom
import openpyxl
import re
import copy

"""
Read in TurningPoint file, course ID and file with student names and return a spreadsheet with the students
on the course and present or absent based on the clicker ID data

input: TurningPOint file name, courseID, spreadsheet with ID's
output: spreadsheet with names and present/absent
"""
def ClickeRegister(turningPointFileName, courseID, studentFile = "/Volumes/Schools/EPMS/HoS/School_Management/Student/27-1-20/Clicker Data CT Fixed (2).xlsx") : 

    # extract TTSession.xml
    tpFile = _ZipFile(turningPointFileName)
    tpFile.extract("TTSession.xml")

    # read TTSession.xml file
    ttFile = open("TTSession.xml")
    
    xmldoc = _minidom.parseString(ttFile.readline())

    # get click device IDS 
    devices = [] 
    iParticipants = 0
    for child in xmldoc.getElementsByTagName("participant") : 
        # print child.childNodes[0].localName,child.childNodes[0].childNodes[0].nodeValue

        deviceID = child.childNodes[0].childNodes[0].nodeValue        
        devices.append(deviceID)

        iParticipants += 1 
        
    print "Found ", iParticipants
    
#
     

"""
 Return string with date of file name string
 Assume filename is of form dd-mm-yy hh-mm.tpxz
 input: filename
 output: dd-mm-yyyy
"""
def getDate(fName):
    import re
    match = re.match(r'^\d\d-\d\d-\d\d\d\d',fName)
    if match:
        return fName[:10]
    else:
        raise Exception("getDate: filename should start with dd-mm-yyyy")
        
        
"""
 Return dict with student data stored in a spreadsheet
 in particular 
     keys are student ID's
     each value is a dict with keys of First Name, Surname, Courses (array)
 input: filename of spreadsheet
 output: dict as describe above    

"""        
def buildStudentDict(filename):
    
# first worksheet has the useful data
    studentFile = openpyxl.load_workbook(filename,data_only=True)    
    ws = studentFile.worksheets[0]
    
    studentData = {}
    for row in ws.iter_rows(max_col=7,values_only=True):
        key = row[2]
        if row[0] is not None:
            if not key in studentData.keys() :
                studentData[key] = {}
                studentData[key]['First'] = row[4]
                studentData[key]['Surname'] = row[6]
                studentData[key]['clicker'] = row[3]
                studentData[key]['courses'] = [row[0] + str(row[1]),]
            else:
                studentData[key]['courses'].append(row[0] + str(row[1]))     
      
            
    return(studentData)

"""
Return subset of studentData where a student is attending a specific course

input: studentData (dict), course (string of from CS|IYNNNN)
output: list with keys of tudentData that has courses matching the course
"""

def selectCourseStudentData(data,course):
   
    attendingCourse = []
    for k,v in data.items():
        if course in v['courses']:
            attendingCourse.append(k)
            
    return(attendingCourse)    

"""
Return subset of studentData with a list of specific clicker ID's

input: studentData (dict), IDs (list)
output: list with keys of studentData that has clicker ID's matching the list of IDs
"""

def selectIDStudentData(data,IDs):
   
    studentIDs = []
    for k,v in data.items():
        if v['clicker'] in IDs:
            print(k)
            studentIDs.append(k)
            
    return(studentIDs)     
        
        
    
"""
Create a dict with studentData based on two lists; 1 list being students registered on a course; another
being students who attended the lecture
New dict has two news values for each key 
1 key is Present (Yes for present, no for absent)
1 key registered on course

"""    

def collateStudentsInLecture(allStudentData,registered,attended):
    r = set(registered)
    a = set(attended)
       
    rUa = r.union(a)
    thisStudentData = {}
    
    for i in rUa:
        thisStudentData[i] = copy.deepcopy(allStudentData[i])
        thisStudentData[i]['Present'] = i in attended
        thisStudentData[i]['Registered'] = i in registered
    
    
    return(thisStudentData)


"""
Create a spreadsheet with present/absent data


"""

def createAttendanceSpreadsheet(collated):
    wb = openpyxl.Workbook()
    ws = wb.create_sheet("Attendance")
    names=("First name","Surname","Student ID","Clicker ID","Present",)
    ws.append(names)
    for k,v in collated.items():
        if v['Registered']:
            if v["Present"]:
                p = "Yes"
            else:
                p = "No"
            thisRow = (v['First'], v['Surname'], k, v['clicker'],p)
            ws.append(thisRow)
    thisRow = ("Students not registered but attended.")
    ws.append(thisRow)
    for k,v in collated.items():
        if not v['Registered']:
            if v["Present"]:
                p = "Yes"
            else:
                p = "No"
            thisRow = (v['First'], v['Surname'], k, v['clicker'],p)
            ws.append(thisRow)
            
            
    wb.save("Test.xlsx")
    
        
    
    
    
