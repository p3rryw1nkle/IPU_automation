from importlib.util import spec_from_loader
import openpyxl

class GetData:
    def get_data(self):
        path = "spreadsheets\Licenses.xlsx"

        # To open the workbook
        # workbook object is created
        wb_obj = openpyxl.load_workbook(path)
        
        # Get workbook active sheet object
        # from the active attribute
        sheet_obj = wb_obj.active

        m_row = 840
        m_col = 25

        myRows = []
        myInitials = "JR"

        for i in range(2, m_row + 1):
            
            rowVals = []
            cell_obj = sheet_obj.cell(row = i, column = 4)

            # if cell_obj.value == myInitials:

            rowVals.append(sheet_obj.cell(row = i, column = 2).value)  # customer name
            rowVals.append(sheet_obj.cell(row = i, column = 12).value) # license number
            rowVals.append(sheet_obj.cell(row = i, column = 19).value) # 2.0 part
            rowVals.append(sheet_obj.cell(row = i, column = 20).value) # long description
            rowVals.append(sheet_obj.cell(row = i, column = 13).value) # expiration date
            rowVals.append(sheet_obj.cell(row = i, column = 10).value) # quantity
            rowVals.append(sheet_obj.cell(row = i, column = 22).value) # city
            rowVals.append(sheet_obj.cell(row = i, column = 23).value) # state
            rowVals.append(sheet_obj.cell(row = i, column = 25).value) # zip code
            rowVals.append(sheet_obj.cell(row = i, column = 26).value) # email address
            rowVals.append(sheet_obj.cell(row = i, column = 27).value) # full name
            rowVals.append(sheet_obj.cell(row = i, column = 24).value) # country

        return myRows
        # for row in myRows:
        #     print(row)


if __name__ == "__main__":
    ss = GetData()

    dr = ss.get_data()

    # for row in dr:
    #     print(row)
