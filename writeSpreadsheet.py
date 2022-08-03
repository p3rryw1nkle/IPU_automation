from datetime import datetime, timedelta
from asyncio.windows_events import NULL
from countryCodes import country_codes
from readSpreadsheet import GetData
import openpyxl
import shutil
import re

class WriteData:
    def create_new_file(self, company, initials, dictionary):

        formatted_name = company.replace("/", "").replace("\\", "") # if there are slashes in the company name, take them out so we can name the IPU file properly

        name = input(f"Company name {company}, please enter a nickname: ") # ask for a company nickname from script runner to put on the IPU form
        default_email = "None"

        sub_list = [i for i in dictionary[company]['email'] if not isinstance(i, int)] # gets a list of only the valid emails
        if len(sub_list) > 0: # if there are valid emails, then put the IPU in the spreadsheets/completed/email folder
            shutil.copyfile("spreadsheets\IPU.xlsx", f"spreadsheets\completed\email\IPU-Clar2.0-{formatted_name}-{initials}.xlsx") # copy the IPU template, rename it, and put it in the proper folder
            path = f"spreadsheets\completed\email\IPU-Clar2.0-{formatted_name}-{initials}.xlsx" # change the file path to access the new blank IPU
            default_email = sub_list[0] # by default, the email that will be used is the first one read in from the company
        else: # if there are no emails found, copy the IPU template and rename it but put it in the 'without_email' folder
            shutil.copyfile("spreadsheets\IPU.xlsx", f"spreadsheets\completed\without_email\IPU-Clar2.0-{formatted_name}-{initials}.xlsx")
            path = f"spreadsheets\completed\without_email\IPU-Clar2.0-{formatted_name}-{initials}.xlsx"

        wb_obj = openpyxl.load_workbook(path) # load the blank IPU template
        sheet = wb_obj.active

        sheet.cell(row=3, column=2).value += name # change the IPU code to user provided company nickname

        # find and change the country code
        country_code = ""
        for code in country_codes: # this uses a dictionary of country codes, which you can find in 'countryCodes.py'
            if dictionary[company]['country'].lower() in country_codes[code]: # if the country is in the dictionary, use the corresponding GL string
                country_code = code
        if country_code == "":
            print(f"Error finding country code for country {dictionary[company]['country']}") # if it cannot find the country code, output an error
        sheet.cell(row=6, column=3).value += country_code

        # for each product license that is stored with the associated company
        for i in range(0, len(dictionary[company]['product id'])):
            count = i
            sheet.cell(row=19 + count, column=1).value = dictionary[company]['product id'][count] # product number (clarity 2.0 part)
            sheet.cell(row=19 + count, column=2).value = dictionary[company]['long description'][count] # product description
            sheet.cell(row=19 + count, column=3).value = dictionary[company]['quantity'][count] # quantity
            sheet.cell(row=19 + count, column=4).value = 0 # cost
            sheet.cell(row=19 + count, column=5).value = 0 # cost
            date = datetime.strftime(dictionary[company]['expiration date'][count], "%m/%d/%Y") # get expiration date
            # change term/dates so that its a week from today until expiration date
            sheet.cell(row=19 + count, column=6).value = f"{(datetime.now()+timedelta(days=7)).month}/{(datetime.now()+timedelta(days=7)).day}/{(datetime.now()+timedelta(days=7)).year} to {date}"

        # customer information
        sheet.cell(row=36, column=2).value = company # full company name
        sheet.cell(row=37, column=2).value = dictionary[company]['address'] # address
        sheet.cell(row=39, column=2).value = f'{str(dictionary[company]["city"])} / {str(dictionary[company]["state"])} / {str(dictionary[company]["zip code"])}' # city / state / zipcode
        sheet.cell(row=40, column=2).value = dictionary[company]['country'] # country
        sheet.cell(row=41, column=2).value = dictionary[company]['contact name'] # primary contact name
        sheet.cell(row=43, column=2).value = default_email # email

        # switch sheet to 'Notes for Operations ONLY'
        for sheet2 in wb_obj:
            if "Notes" in sheet2.title:
                sheet = sheet2

        # append license numbers
        licenses = ""
        for license in dictionary[company]['license']:
            licenses += license + ", " # each license is separated by commas
        licenses = licenses[0:-2] # gets rid of the last ', ' at the end
        sheet.cell(row=3, column=3).value += licenses # add licenses to cell 'Original Lic#:'

        wb_obj.save(path) # save the IPU


    def process_files(self, initials):
        myData = GetData(initials=initials) # uses 'readSpreadsheet' to get your assigned IPU data
        data_dict = myData.get_data()
        myData.check_validity() # checks to make sure emails are valid. If not, it will still create the IPUs but it will log conflicting
                                # emails to logs/conflicts.log

        for company in data_dict.keys(): # for each company in the data dictionary
            self.create_new_file(company, initials=initials, dictionary=data_dict) # create a new IPU file

if __name__ == "__main__":
    makeFiles = WriteData()
    makeFiles.process_files(initials="LB") # put your initials here