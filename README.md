# contactCsvCreator
contactCsvCreator is a tool that parses contact details from a Google sheet and creates a CSV file. This CSV file can then be used to directly import contacts using the import option in [Google Contacts](https://contacts.google.com/).

## But, how does it benefit me?
Imagine a scenario : You're going to kick-start a new course and have circulated a Google Form for students who are interested. The form responses are over a 1000 entries, with each containing the Name, Number, Email, DOB and some other details of the students. You now want to frequently communicate with these students over Whatsapp/Email or any other means, which requires creating a contact for each entry, and filling it in with the provided details. *What do you do? You use contactCsvCreator.*

## Requirements
1. Run `pip3 install -r requirements.txt`.
2. This script internally uses the [gspread](https://docs.gspread.org/en/v5.7.0/) python library to parse the data present in the Google Sheet. For accessing the sheet, the user/service account needs to authorize using OAuth Client API present in the library.
3. To create a service account with the intended access rights, refer the configuration steps documented here - https://docs.gspread.org/en/v5.7.0/oauth2.html#for-bots-using-service-account. If all of them are followed correctly, the file containing the service account credentials should be available for download as a Json file. As part of the API requirement, this file needs to be stored as : **~/.config/gspread/service_account.json** in your local system.

## Usage
1. Simply run the command : `python createContactFile.py`
2. A CSV file will be created in the **./outputCsv/** directory. It's name is created according to the convention : `Contacts - <Spreadsheet Name> - <Sheet Name>.csv`
3. Import this CSV file to [Google Contacts](https://contacts.google.com/)
4. It is now seen that contacts have been created with the respective details that are present in their corresponding entry in the Google Sheet. These can now be directly used for communication, maintenance of records, and alike.

### The script asks for the following inputs :
1. Spreadsheet Name : This is the parent spreadsheet which contains multiple sheets within it.
2. Sheet Name : This sheet that is present in the spreadsheet, and is given READ/WRITE access to the service account
3. Addition of mappings : The columns of the Google Sheet are mapped to the columns of the CSV sheet. We can specify mappings that are other than the default mapping, which is -
```
+-------+-----------------------------+---------------+-------------------+
|   idx | Name                        | Family Name   | Phone 1 - Value   |
+=======+=============================+===============+===================+
|     1 | ['First Name', 'Last Name'] | Last Name     | Mobile            |
+-------+-----------------------------+---------------+-------------------+

```

## Scenarios
### Usecase 1
Without adding any mapping :

[SEE THE USAGE LIVE IN ACTION](https://user-images.githubusercontent.com/108089086/202208266-fad42e82-6dbd-4e1c-a882-5934011aa296.webm)


### Usecase 2
Adding a new mapping :

[SEE THE USAGE LIVE IN ACTION](https://user-images.githubusercontent.com/108089086/202204796-9dc3bdfb-18af-452e-b9fa-3dd996d66388.webm)
