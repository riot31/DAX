import ftplib
import xml.etree.ElementTree as ET
from os import listdir
from os.path import isfile, join


PATH   =  'C:\\Users\\jbUser\\Desktop\\DAX\\'
TARGET_FILE = 'C:\\Users\\jbUser\\Desktop\\result.csv'
TARGET = '{http://schemas.microsoft.com/dynamics/2008/01/documents/SalesForceCust}'
BAD_KEYS  = set(['CustTable_1', 'SalesForceCust'])
KYES_FOR_BOOL_CONVERT = set(['rjvOptSendCatalog', 'rjvOptSendEmail', 'rjvOptShareInfo'])
SEPARATOR = ';'
LASTNAME_KEY = 'jsLastName'
EMAIL_KEY = 'Email'
RECORDTYPE_KEY = 'RecordTypeLabel'
DAX_STR2_KEY = 'jsStreet2'
DAX_STR1_KEY = 'jsStreet1'
DAX_STR3_KEY = 'jsStreet3'
SF_STR_KEY = 'Account_BillingStreet_c'

def convertStreet(str1, str2, str3):
    strs = []
    if(type(str1) != type(None)):
        strs.append(str1)
    if(type(str2) != type(None)):        
        strs.append(str2)
    if(type(str3) != type(None)):
        strs.append(str3)
    return '\n'.join(strs)

def checkStreet(csvMap):
    str1 = csvMap.get(DAX_STR1_KEY)
    str2 = csvMap.get(DAX_STR2_KEY)
    str3 = csvMap.get(DAX_STR3_KEY)
    csvMap.update({SF_STR_KEY: convertStreet(str1, str2, str3)})

def convertBool(value):
    return value == 'Yes'

def convertRecordType(lastName, email):
    LastName = ''
    if(lastName  != ''):
        LastName = lastName;
    if(lastName == '' and email != ''):
        LastName = 'Unknown';
    if(LastName == ''):
        recordType = 'Account';
    else:
        recordType = 'Contact';
    return recordType, LastName

def checkBool(csvMap):
    for key in KYES_FOR_BOOL_CONVERT :
        value = csvMap.get(key)
        if(value != None):
            csvMap.update({key: convertBool(value)})

def checkRecordType(csvMap):
    email = csvMap.get(EMAIL_KEY)
    lastName = csvMap.get(LASTNAME_KEY)
    recordType, LastName = convertRecordType(lastName, email)
    csvMap.update({RECORDTYPE_KEY: recordType})
    csvMap.update({LASTNAME_KEY: LastName})

def converting(csvMap):
    checkBool(csvMap)
    checkRecordType(csvMap)
    checkStreet(csvMap)
    

columns = set();
rows = [];
csvMap = {}
for file in [f for f in listdir(PATH) if isfile(join(PATH, f))]:
    tree = ET.parse(PATH + file)
    csvMap = {}
    for node in tree.iter():
        if(node.tag.startswith(TARGET)):
            key = node.tag.replace(TARGET, '')
            if(not key in BAD_KEYS):
                columns.add(key)
                csvMap.update({key: node.text})
    converting(csvMap)
    rows.append(csvMap)


columns.add(RECORDTYPE_KEY)     
columns.add(SF_STR_KEY)                
file = open(TARGET_FILE,'w') 
textRow = ''
for key in columns:
    textRow += key + SEPARATOR;
file.write(textRow[:len(textRow) - 1]+ '\n') 



for row in rows:
    textRow = ''
    for key in columns:
        value = row.get(key)
        if(value == None):
            value = ''
        textRow += str(value) + SEPARATOR;
    file.write(textRow[:len(textRow) - 1] + '\n') 

file.close()



