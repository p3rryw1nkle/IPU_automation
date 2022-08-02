from asyncio.windows_events import NULL
from readSpreadsheet import GetData
from countryCodes import countryCodes
import openpyxl
import shutil
import datetime

class WriteData:

    def create_new_file(row):

        # make a copy of IPU template
        # if there is an email, put it into the email folder
        # otherwise, put it into without email folder
        # replace fields with information from row
        # save

        name = input(f"Company name {row[0]}, please enter a nickname: ")

        if row[9] not in [NULL, 0]: # if there is an email
            shutil.copyfile("spreadsheets\IPU.xlsx", f"spreadsheets\completed\email\IPU-Clar2.0-{name}-LB.xlsx")

            path = f"spreadsheets\completed\email\IPU-Clar2.0-{name}-LB.xlsx"
        else:
            shutil.copyfile("spreadsheets\IPU.xlsx", f"spreadsheets\completed\without_email\IPU-Clar2.0-{name}-LB.xlsx")

            path = f"spreadsheets\completed\without_email\IPU-Clar2.0-{name}-LB.xlsx"

        wb_obj = openpyxl.load_workbook(path)
        sheet = wb_obj.active

        sheet.cell(row=3, column=2).value += name # company name

        # write country code
        country_code = ""

        for code in countryCodes:
            # print(code)
            if row[11].lower() in code:
                country_code = code
        
        sheet.cell(row=6, column=3).value += country_code

        sheet.cell(row=19, column=1).value = row[2] # product number
        sheet.cell(row=19, column=2).value = row[3] # product description
        sheet.cell(row=19, column=3).value = row[5] # quantity
        

        date = datetime.datetime.strftime(row[4], "%m/%d/%Y")

        sheet.cell(row=19, column=6).value = "8/8/2022 to " + date # quantity

        # customer information
        sheet.cell(row=36, column=2).value = row[0] # name 
        sheet.cell(row=37, column=2).value = row[12] # address
        sheet.cell(row=39, column=2).value = str(row[6]) + " " + str(row[7]) + " " + str(row[8]) # city + state + zipcode
        sheet.cell(row=40, column=2).value = row[11] # country
        sheet.cell(row=41, column=2).value = row[10] # email

        for sheet2 in wb_obj:
            if "Notes" in sheet2.title:
                sheet = sheet2

        sheet.cell(row=3, column=3).value += row[1]

        wb_obj.save(path)


    def process_files(self):
        myData = GetData(initials="LB")
        myData = myData.get_data()

        print(len(myData))

        for row in myData:
            self.create_new_file(row)

if __name__ == "__main__":
    makeFiles = WriteData
    makeFiles.process_files(makeFiles)