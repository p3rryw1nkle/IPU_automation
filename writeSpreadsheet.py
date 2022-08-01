from venv import create
from readSpreadsheet import GetData
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

        print(row[9])

        if row[9]:
            shutil.copyfile("spreadsheets\IPU.xlsx", f"spreadsheets\completed\email\IPU-Clar2.0-{row[0]}-LB.xlsx")

            path = f"spreadsheets\completed\email\IPU-Clar2.0-{row[0]}-LB.xlsx"
        else:
            shutil.copyfile("spreadsheets\IPU.xlsx", f"spreadsheets\completed\without_email\IPU-Clar2.0-{row[0]}-LB.xlsx")

            path = f"spreadsheets\completed\without_email\IPU-Clar2.0-{row[0]}-LB.xlsx"

        wb_obj = openpyxl.load_workbook(path)
        sheet = wb_obj.active

        name = sheet.cell(row=3, column=2).value
        name += row[0][0:5]
        sheet.cell(row=3, column=2).value = name

        sheet.cell(row=19, column=1).value = row[2] # product number
        sheet.cell(row=19, column=2).value = row[3] # product description
        sheet.cell(row=19, column=3).value = row[5] # quantity

        date = datetime.datetime.strftime(row[4], "%m/%d/%Y")

        sheet.cell(row=19, column=6).value = "8/8/2022 to " + date # quantity

        sheet.cell(row=36, column=2).value = name
        sheet.cell(row=39, column=2).value = str(row[6]) + " " + str(row[7]) + " " + str(row[8])
        sheet.cell(row=41, column=2).value = row[10]

        wb_obj.save(path)


    def process_files(self):
        myData = GetData
        myData = myData.get_data()

        self.create_new_file(myData[0])

        # for row in myData:
        #     self.create_new_file(row)


    # make a copy of excel spreadsheet
    # write to specific cells
    # save to completed file

if __name__ == "__main__":
    makeFiles = WriteData
    makeFiles.process_files(makeFiles)