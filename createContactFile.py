import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pprint
import tabulate
import csv
import sys
import string
from os import path, mkdir
from colors import *

# Output Directory Name
OUTPUTDIR = "outputCsv/"

# Final attributes to be used in the output CSV file.
contactFileAttributes     = ['Name', 'Given Name', 'Additional Name', 'Family Name', 'Yomi Name', 'Given Name Yomi', 'Additional Name Yomi', 'Family Name Yomi', 'Name Prefix', 'Name Suffix', 'Initials', 'Nickname', 'Short Name', 'Maiden Name', 'Birthday', 'Gender', 'Location', 'Billing Information', 'Directory Server', 'Mileage', 'Occupation', 'Hobby', 'Sensitivity', 'Priority', 'Subject', 'Notes', 'Language', 'Photo', 'Group Membership', 'E-mail 1 - Type', 'E-mail 1 - Value', 'Phone 1 - Type', 'Phone 1 - Value']

# Must have attributes for the output CSV file.
contactRequiredAttributes = [ "Name", "Phone 1 - Value"]

# Google Sheet attribute list
sheetAllAttributes = ["First Name", "Last Name", "Email Address", "Mobile", "City", "Occupation"]

# Must have attributes in the Google Sheet
sheetRequiredAttributes = ["First Name", "Last Name", "Mobile"]


MappingsSheetToCsv = {
"Name" : ["First Name" , "Last Name"],
"Family Name" : "Last Name",
"E-mail 1 - Value" : "Email Address",
"Phone 1 - Value" : "Mobile",
"Occupation" : "Occupation",
"Location" : "City"
}

def displayTable(dict_list):
    """
    This function
        1. Takes a list of dictionary
        2. Add an Index column, and
        3. Displays the data in tabular format
    """
    header = ['idx'] + list(dict_list[0].keys())
    rows = [[idx + 1] + list(x.values()) for idx, x in enumerate(dict_list)]
    print(tabulate.tabulate(rows, header, tablefmt='grid'))


def displayTableFromList(listname, list_vals):
    tmp_dict = []
    for temp in list_vals:
        tmp = {listname : temp }
        tmp_dict.append(tmp)

    displayTable(tmp_dict)


def selectOptionFromList(listname, list_vals):
    displayTableFromList(listname, list_vals)
    index = int(input('Enter {} index: '.format(listname)))
    if index > len(list_vals):
        prRed("Wrong choice entered, please try again!")
        sys.exit()

    optionVal = list_vals[index - 1]
    prBlue("The {} Selected is \"{}\"".format(listname, optionVal))
    return index, optionVal



class Mappings:
    def __init__(self, idealMap, dList, sList):
        self.idealMapping   = idealMap
        self.dstMappingList = dList
        self.srcMappingList = sList

    def createFinalMappingBasedOnFiles(self):
        self.finalMapping = {}
        for key, value in self.idealMapping.items():
            if type(value) == list:
                flag = all(item in self.srcMappingList for item in value)
            else:
                flag = value in self.srcMappingList
            if flag and key in self.dstMappingList:
                self.finalMapping[key] = value

    def updateMapping(self):
        displayDictionaryAsMappings(self.finalMapping)

        choice = input("\nDo you want to add any more mappings? Type yes or no: ")
        while choice == "yes":
            fieldIndex1, fieldValue1 = selectOptionFromList('Destination File Field', self.dstMappingList)
            fieldIndex2, fieldValue2 = selectOptionFromList('Source File Field', self.srcMappingList)
            self.finalMapping[fieldValue1] = fieldValue2
            prYellow('\nCurrent Mappings : ')
            displayDictionaryAsMappings(self.finalMapping)
            choice = input("\nDo you want to add any more mappings? Type yes or no: ")
        else:
            prGreen("\n[+] Final Mappings")

        displayDictionaryAsMappings(self.finalMapping)

class RecordClass:
    def __init__(self, rcList, rcAttr):
        self.allRecords = rcList
        self.recordAttributes = rcAttr
        self.recordSize = len(self.allRecords)
        self.trimmedRecords = []

    def checkRequiredAttributes(self):
        missingFields = []

        if self.bIsDestFile:
            for field in self.allAttributes:
                if field not in self.recordAttributes:
                    missingFields.append(field)
        else:
            for field in self.requiredAttributes:
                if field not in self.recordAttributes:
                    missingFields.append(field)

        if (len(missingFields)):
            prRed("\nThere are missing fields in the sheet, they are : ")
            displayTableFromList('Field Name', missingFields)
            sys.exit()


    def fetchAttrValueListFromRecords(self, attrName):
        contactNumbers = list(set([ num[attrName] for num in self.allRecords ]))
        return contactNumbers


    def setAllAttributes(self, allAttr, requiredAttr):
        self.allAttributes      = allAttr
        self.requiredAttributes = requiredAttr

    def removeDuplicatesBasedOnValue(self, uniqVal=None):
        new_list = []
        new_list_vals = []
        if uniqVal is not None:
            for item in self.allRecords:
                if item[uniqVal] not in new_list_vals and item[uniqVal] != "None":
                    new_list.append(item)
                    new_list_vals.append(item[uniqVal])

            # Update the records and record size as well
            prLightPurple("[!] Number of Removed Duplicates : {}".format(len(self.allRecords) - len(new_list)))
            self.allRecords = new_list
            self.recordSize = len(self.allRecords)


class FileClass(RecordClass):
    def __init__(self, name1 = "", name2 = "", ext=""):
        self.fileName = OUTPUTDIR + "Contacts - " + name1 + " - " + name2 + ext
        if path.isfile(self.fileName):
            self.bFileExists = True
        else:
            self.bFileExists = False
        self.bIsDestFile = False


    def fetchRecordsFromFile(self):
        if not self.bFileExists:
            prRed("Cant fetch records from a file that doesn't exist! Returning...")
            sys.exit()

        self.fileRecords = []
        with open(self.fileName, 'r') as csFile:
            reader = csv.DictReader(csFile)
            for dataRow in reader:
                self.fileRecords.append(dataRow)

        self.fileRecordFields  = list(self.fileRecords[0].keys())
        super().__init__(self.fileRecords, self.fileRecordFields)



class SheetClass(RecordClass):
    def __init__(self, sheetClnt, name1 = "", name2 = ""):
        try :
            sheetClnt.open(name1).worksheet(name2)
        except :
            prRed("\nPlease check if you've inputted correct values!\nSpreadsheet Name : " + name1 + "\nSheet Name : " + name2)
            sys.exit()

        self.spreadsheetName = name1
        self.sheetName       = name2
        self.sheetObject     = sheetClnt.open(self.spreadsheetName).worksheet(self.sheetName)
        self.sheetHeaders    = self.sheetObject.row_values(1)
        self.bIsDestFile     = False

        super().__init__(self.sheetObject.get_all_records(default_blank='None'), self.sheetHeaders)


    def fetchColumnIndex(self):
        displayTableFromList('Column', self.sheetHeaders)
        index = int(input('Enter Column index: '))
        return index


    def fetchColumnValues(self, columnIndex=None):

        if columnIndex is None:
            columnIndex = self.fetchColumnIndex()

        # Perform a sanity check
        if columnIndex > len(self.sheetHeaders):
            prRed("You have selected an invalid column name : {}".format(columnIndex));
            sys.exit()

        columnValues = self.sheetObject.col_values(columnIndex)
        columnValues.remove(self.sheetHeaders[columnIndex - 1])
        columnValues = list(set(columnValues))

        return columnValues, self.sheetHeaders[columnIndex - 1]


def getWorksheetName():
  spreadsheetName = str(input("Enter the spreadsheet name: "))
  spreadsheetName = spreadsheetName.strip()
  sheetName = str(input("Enter the Sheet name: "))
  sheetName = sheetName.strip()
  if spreadsheetName and sheetName:
    prYellow ("\nThe Spreadsheet selected is : " + spreadsheetName + "\nThe Sheet selected is : " + sheetName + "\n")
  else:
    prYellow("Please enter the name of the Spreadsheet and Sheet!");
    exit();
  return spreadsheetName, sheetName;


def getAuthClient():
    # Scope of our spreadsheet
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

    # Get the credentials
    creds = ServiceAccountCredentials.from_json_keyfile_name('Customer_Database_secret.json', scope)

    # Authorise the client
    return gspread.authorize(creds)


def checkOrCreateDestDir(dirName):
    if not path.exists(dirName):
        prGreen("\n\n[+] Creating {} directory in the same folder!".format(OUTPUTDIR))
        mkdir(dirName)


def fetchTheClientRecords(allAttr=None, reqAttr=None):

    # Authorise the client
    client = getAuthClient()

    # Get the spreadsheet and Sheet Name
    spreadsheetName, sheetName = getWorksheetName()

    prGreen("[+] Fetching all the records...")
    clientSheet = SheetClass(client, spreadsheetName, sheetName)

    # Check if the following headers are present in the sheet, otherwise exit
    if allAttr == None and reqAttr == None:
        clientSheet.setAllAttributes(sheetAllAttributes, sheetRequiredAttributes)
    else:
        clientSheet.setAllAttributes(allAttr, reqAttr)

    clientSheet.checkRequiredAttributes()
    return clientSheet


def createDictionaryFromMappings(mapping, contactAttributes, recordVal):
    tmpDictionary = {}

    for attr in contactAttributes:
        if attr in mapping.keys():
            if type(mapping[attr]) == list:
                tmpDictionary[attr] = ""
                for index in mapping[attr]:
                    tmpDictionary[attr] += recordVal[index] + " "
                tmpDictionary[attr] = str(tmpDictionary[attr].strip())
            else:
                tmpDictionary[attr] = str(recordVal[mapping[attr]])
        else:
            tmpDictionary[attr] = ""

    return tmpDictionary


def removeExistingContacts(clientContactsFileObject, clientSheetObject):

    # Note that these are in string format
    currentSavedContacts = clientContactsFileObject.fetchAttrValueListFromRecords('Phone 1 - Value')

    clientSheetObject.trimmedRecords = []
    for i in range(clientSheetObject.recordSize):
        mobile = str(clientSheetObject.allRecords[i]['Mobile'])
        if mobile not in currentSavedContacts and mobile != "None":
            clientSheetObject.trimmedRecords.append(clientSheetObject.allRecords[i])


    if len(clientSheetObject.trimmedRecords) == 0:
        prBlue("[++] There are no new contacts.. Exiting!")
        sys.exit()

    prLightPurple("[!] Number of new contacts : {}".format(len(clientSheetObject.trimmedRecords)))


def createContactFileObject(Prefix, Suffix, bIsDest, fileExt=""):
    contactFileObject = FileClass(Prefix, Suffix, fileExt)
    contactFileObject.bIsDestFile = bIsDest
    contactFileObject.setAllAttributes(contactFileAttributes, contactRequiredAttributes)

    return contactFileObject


def getRecordsWithAttr(recordList, attrKey, attrValue):
    if len(recordList) <= 0:
        prRed("Record List is empty!")
        sys.exit()

    final_record = []
    for client_deet in recordList:
        if client_deet.get(attrKey) == attrValue:
            final_record.append(client_deet)

    return final_record


def displayKeyValuePairs(dictName):
    prYellow('This is how the mapping is done :')
    for key, value in dictName.items():
        if type(value) == list:
            val = " + ".join(value)
            print (f"{key:<30}{'==>':<10}{val:<40}")
        else:
            print (f"{key:<30}{'==>':<10}{value:<40}")


def displayDictionaryAsMappings(dictName):
    listDict = []
    listDict.append(dictName)
    displayTable(listDict)


def addRecordsToFile(fileName, fileHeaders, operationChoice, records):

    fileName = fileName
    if operationChoice == "append":
        with open(fileName, 'a') as csvFile:
            writer = csv.DictWriter(csvFile, fieldnames = fileHeaders)
            writer.writerows(records)

    elif operationChoice == "write":
        with open(fileName, 'w') as csvFile:
            writer = csv.DictWriter(csvFile, fieldnames = fileHeaders)
            writer.writeheader()
            writer.writerows(records)
    else:
        prRed("Unknown Choice... Exiting!")
        sys.exit()

    prBlue("[++] Contacts added to File : {}\n".format(fileName))


def createDestinationFileFromSourceRecords(objMapping, destFileObject, sourceRecords, bIsDstFileTmp):
        recordFinalList = []
        for record in sourceRecords:
            recordFinalList.append(createDictionaryFromMappings(objMapping.finalMapping, objMapping.dstMappingList ,record))

        fileName    = destFileObject.fileName
        fileHeaders = objMapping.dstMappingList

        if destFileObject.bFileExists and not bIsDstFileTmp:
            addRecordsToFile(fileName, fileHeaders, "append", recordFinalList)
        else:
            addRecordsToFile(fileName, fileHeaders, "write", recordFinalList)


def main():

    # Fetch the client records from the spreadsheet
    clientSheetObject = fetchTheClientRecords()
    prLightPurple("[!] Number of Fetched Records : {}".format(len(clientSheetObject.allRecords)))

    prGreen("[+] Removing all the duplicated and invalid rows...")
    clientSheetObject.removeDuplicatesBasedOnValue('Mobile')

    # Create a contact file object
    clientContactsFileObject = createContactFileObject(
        clientSheetObject.spreadsheetName,
        clientSheetObject.sheetName,
        True,
        ".csv"
    )

    currentContactTmpObject = createContactFileObject(
        "My",
        "Temp",
        True,
        ".csv"
    )

    if clientContactsFileObject.bFileExists:
        # Fetch all the records
        prLightGray("\n[!] There is already an existing file for this sheet, fetching its Records...")
        clientContactsFileObject.fetchRecordsFromFile()

        # Check for required Fields
        clientContactsFileObject.checkRequiredAttributes()

        # If the file has some records, remove them from the fetched records
        prGreen("[+] Removing all the Records that are already present in the file...")
        removeExistingContacts(clientContactsFileObject, clientSheetObject)


    prYellow("\n\nThe current Mapping is as follows :")
    mappingSheetToCsv = Mappings(MappingsSheetToCsv, clientContactsFileObject.allAttributes, clientSheetObject.recordAttributes)
    mappingSheetToCsv.createFinalMappingBasedOnFiles()
    mappingSheetToCsv.updateMapping()

    if clientContactsFileObject.bFileExists:
        recordList = clientSheetObject.trimmedRecords
    else:
        recordList = clientSheetObject.allRecords

    # Check if directory exists, and if not, create one!
    checkOrCreateDestDir(OUTPUTDIR)

    # Adding to main file
    prGreen("\n[+] Adding all the new contacts to the main file...")
    createDestinationFileFromSourceRecords(mappingSheetToCsv, clientContactsFileObject, recordList, False)

    # Adding to temporary File
    if len(clientSheetObject.trimmedRecords):
        prGreen("\n[+] Adding all the new contacts to the temporary file...")
        createDestinationFileFromSourceRecords(mappingSheetToCsv, currentContactTmpObject, clientSheetObject.trimmedRecords, True)


if __name__ == '__main__':
    main()
