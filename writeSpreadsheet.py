from asyncio.windows_events import NULL
from countryCodes import country_codes as countryCodes
from readSpreadsheet import GetData
import datetime
import openpyxl
import shutil

class WriteData:

    def create_new_file(company, company_name):

        # make a copy of IPU template
        # if there is an email, put it into the email folder
        # otherwise, put it into without email folder
        # replace fields with information from row
        # save

        name = input(f"Company name {company_name}, please enter a nickname: ")
        default_email = "None"

        for email in company['email']:
            if email not in [NULL, 0]: # if there is an email
                shutil.copyfile("spreadsheets\IPU.xlsx", f"spreadsheets\completed\email\IPU-Clar2.0-{name}-LB.xlsx")
                path = f"spreadsheets\completed\email\IPU-Clar2.0-{name}-LB.xlsx"
                default_email = email
                break
            else: # if there is not an email, put in without_email folder
                shutil.copyfile("spreadsheets\IPU.xlsx", f"spreadsheets\completed\without_email\IPU-Clar2.0-{name}-LB.xlsx")
                path = f"spreadsheets\completed\without_email\IPU-Clar2.0-{name}-LB.xlsx"

        wb_obj = openpyxl.load_workbook(path)
        sheet = wb_obj.active

        sheet.cell(row=3, column=2).value += name # company name

        # write country code
        country_code = ""
        for code in countryCodes:
            # print(code)
            if company['country'].lower() in countryCodes[code]:
                country_code = code
        if country_code == "":
            print(f"Error finding country code for country {company['country']}")
        sheet.cell(row=6, column=3).value += country_code

        # write products
        for i in range(0, len(company['product id'])):
            count = i
            sheet.cell(row=19 + count, column=1).value = company['product id'][count] # product number
            sheet.cell(row=19 + count, column=2).value = company['long description'][count] # product description
            sheet.cell(row=19 + count, column=3).value = company['quantity'][count] # quantity
            sheet.cell(row=19 + count, column=4).value = 0 # cost
            sheet.cell(row=19 + count, column=5).value = 0 # cost
            date = datetime.datetime.strftime(company['expiration date'][count], "%m/%d/%Y")
            sheet.cell(row=19 + count, column=6).value = "8/9/2022 to " + date # quantity

        # customer information
        sheet.cell(row=36, column=2).value = company_name # name
        sheet.cell(row=37, column=2).value = company['address'] # address
        sheet.cell(row=39, column=2).value = str(company['city']) + " " + str(company['state']) + " " + str(company['zip code']) # city + state + zipcode
        sheet.cell(row=40, column=2).value = company['country'] # country
        sheet.cell(row=41, column=2).value = company['contact name'] # contact name
        sheet.cell(row=43, column=2).value = default_email # email

        # switch sheets
        for sheet2 in wb_obj:
            if "Notes" in sheet2.title:
                sheet = sheet2

        # append license number
        licenses = ""
        for license in company['license']:
            licenses += license + ", "
        licenses = licenses[0:-2]

        sheet.cell(row=3, column=3).value += licenses

        wb_obj.save(path)


    def process_files(self):
        myData = GetData(initials="LB")
        myData = myData.get_data()

        print(len(myData))

        for company in myData:
            self.create_new_file(myData[company], company)

if __name__ == "__main__":
    makeFiles = WriteData
    makeFiles.process_files(makeFiles)