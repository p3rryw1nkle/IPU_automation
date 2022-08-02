from asyncio.windows_events import NULL
from countryCodes import country_codes as countryCodes
from readSpreadsheet import GetData
import datetime
import openpyxl
import shutil
import re

class WriteData:
    def create_new_file(self, company, initials, dictionary):

        # make a copy of IPU template
        # if there is an email, put it into the email folder
        # otherwise, put it into without email folder
        # replace fields with information from row
        # save

        formatted_name = company.replace("/", "").replace("\\", "")

        name = input(f"Company name {company}, please enter a nickname: ")
        default_email = "None"

        sub_list = [i for i in dictionary[company]['email'] if not isinstance(i, int)]
        if len(sub_list) > 0:
            shutil.copyfile("spreadsheets\IPU.xlsx", f"spreadsheets\completed\email\IPU-Clar2.0-{formatted_name}-{initials}.xlsx")
            path = f"spreadsheets\completed\email\IPU-Clar2.0-{formatted_name}-{initials}.xlsx"
            default_email = sub_list[0]
        else:
            shutil.copyfile("spreadsheets\IPU.xlsx", f"spreadsheets\completed\without_email\IPU-Clar2.0-{formatted_name}-{initials}.xlsx")
            path = f"spreadsheets\completed\without_email\IPU-Clar2.0-{formatted_name}-{initials}.xlsx"

        wb_obj = openpyxl.load_workbook(path)
        sheet = wb_obj.active

        sheet.cell(row=3, column=2).value += name # company name

        # write country code
        country_code = ""
        for code in countryCodes:
            # print(code)
            if dictionary[company]['country'].lower() in countryCodes[code]:
                country_code = code
        if country_code == "":
            print(f"Error finding country code for country {dictionary[company]['country']}")
        sheet.cell(row=6, column=3).value += country_code

        # write products
        for i in range(0, len(dictionary[company]['product id'])):
            count = i
            sheet.cell(row=19 + count, column=1).value = dictionary[company]['product id'][count] # product number
            sheet.cell(row=19 + count, column=2).value = dictionary[company]['long description'][count] # product description
            sheet.cell(row=19 + count, column=3).value = dictionary[company]['quantity'][count] # quantity
            sheet.cell(row=19 + count, column=4).value = 0 # cost
            sheet.cell(row=19 + count, column=5).value = 0 # cost
            date = datetime.datetime.strftime(dictionary[company]['expiration date'][count], "%m/%d/%Y")
            sheet.cell(row=19 + count, column=6).value = "8/9/2022 to " + date # quantity

        # customer information
        sheet.cell(row=36, column=2).value = company # name
        sheet.cell(row=37, column=2).value = dictionary[company]['address'] # address
        sheet.cell(row=39, column=2).value = str(dictionary[company]['city']) + " " + str(dictionary[company]['state']) + " " + str(dictionary[company]['zip code']) # city + state + zipcode
        sheet.cell(row=40, column=2).value = dictionary[company]['country'] # country
        sheet.cell(row=41, column=2).value = dictionary[company]['contact name'] # contact name
        sheet.cell(row=43, column=2).value = default_email # email

        # switch sheets
        for sheet2 in wb_obj:
            if "Notes" in sheet2.title:
                sheet = sheet2

        # append license number
        licenses = ""
        for license in dictionary[company]['license']:
            licenses += license + ", "
        licenses = licenses[0:-2]

        sheet.cell(row=3, column=3).value += licenses

        wb_obj.save(path)


    def process_files(self, initials):
        myData = GetData(initials=initials)
        data_dict = myData.get_data()
        myData.check_validity()

        # print(len(myData))

        for company in data_dict.keys():
            self.create_new_file(company, initials=initials, dictionary=data_dict)

if __name__ == "__main__":
    makeFiles = WriteData()
    makeFiles.process_files(initials="LB")