# contactCsvCreator
contactCsv creator is a tool that parses contact details from a Google sheet and creates a CSV file. This CSV file can then be directly used to import contacts using the Google Contacts import option in https://contacts.google.com/. **This script is most beneficial in times when responses of a Google Form contain identification details of users like their name, numbers, city, occupation, etc. The form owner can directly use this script on the response sheet to create a CSV file, thus enabling them to import all the entries as phone contacts**

## Requirements 
1. Run `pip3 install -r requirements.txt`.
2. This script internally uses the [gspread](https://docs.gspread.org/en/v5.7.0/) python library to parse the data present in the Google Sheet. Hence, the user/service account needs to authorize using OAuth Client API. To configure this, refer to https://docs.gspread.org/en/v5.7.0/oauth2.html#for-bots-using-service-account. If the configuration was success, the file **~/.config/gspread/service_account.json** should be present and contains the service account credentials that will be used by the API.

## Usage
1. Simply run the command : `python createContactFile.py`
2. A CSV file will be created in the **./outputCsv/** directory.
3. Import this CSV file to [Google Contacts](https://contacts.google.com/)
4. You'll now see that the contacts have been created with the respective details that are present in their corresponding entry in the Google Sheet

### The script asks for the following inputs :
1. Spreadsheet Name : This is the parent spreadsheet which contais multiple sheets within it.
2. Sheet Name : This the sheet that is part of the spreadsheet, and is given the access to the service account
3. Addition of mappings : The columns of the Google Sheet are mapped to the columns of the CSV sheet. We can specify mappings that are other than the default mapping, which is :
```
+-------+-----------------------------+---------------+-------------------+
|   idx | Name                        | Family Name   | Phone 1 - Value   |
+=======+=============================+===============+===================+
|     1 | ['First Name', 'Last Name'] | Last Name     | Mobile            |
+-------+-----------------------------+---------------+-------------------+

```

## Scenarios
### Usecase 1
Without adding any mapping

[SEE THE USAGE LIVE IN ACTION](https://user-images.githubusercontent.com/108089086/202208266-fad42e82-6dbd-4e1c-a882-5934011aa296.webm)


### Usecase 2
Adding a new mapping

[SEE THE USAGE LIVE IN ACTION](https://user-images.githubusercontent.com/108089086/202204796-9dc3bdfb-18af-452e-b9fa-3dd996d66388.webm)

