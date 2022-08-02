from importlib.util import spec_from_loader
from pprint import pprint
import openpyxl
import logging


class GetData:
    def __init__(self, initials):
        self.store_dict = {}
        self.row_vals = []
        self.initials = initials

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

        for i in range(2, m_row + 1):

            cell_obj = sheet_obj.cell(row = i, column = 4)

            if cell_obj.value is not None and cell_obj.value.lower() == self.initials.lower():

                company_name = self.append_and_return(sheet_obj, row=i, col=2)
                license_num = self.append_and_return(sheet_obj, row=i, col=12)
                prod_id = self.append_and_return(sheet_obj, row=i, col=19)
                long_desc = self.append_and_return(sheet_obj, row=i, col=20)
                exp_date = self.append_and_return(sheet_obj, row=i, col=13)
                quantity = self.append_and_return(sheet_obj, row=i, col=10)
                phys_address = self.append_and_return(sheet_obj, row=i, col=21)
                city = self.append_and_return(sheet_obj, row=i, col=22)
                state = self.append_and_return(sheet_obj, row=i, col=23)
                zip_code = self.append_and_return(sheet_obj, row=i, col=25)
                email_address = self.append_and_return(sheet_obj, row=i, col=26)
                full_name = self.append_and_return(sheet_obj, row=i, col=27)
                country = self.append_and_return(sheet_obj, row=i, col=24)

                if company_name in self.store_dict:
                    self.store_dict[company_name]['license'].append(license_num)
                    self.store_dict[company_name]['product id'].append(prod_id)
                    self.store_dict[company_name]['long description'].append(long_desc)
                    self.store_dict[company_name]['expiration date'].append(exp_date)
                    self.store_dict[company_name]['quantity'].append(quantity)
                    self.store_dict[company_name]['email'].append(email_address if email_address not in self.store_dict[company_name]['email'] else 0)
                else:
                    self.store_dict[company_name] = {'license': [license_num],
                                                     'product id': [prod_id],
                                                     'long description': [long_desc],
                                                     'expiration date': [exp_date],
                                                     'quantity': [quantity],
                                                     'address': phys_address,
                                                     'city': city,
                                                     'state': state,
                                                     'country': country,
                                                     'zip code': zip_code,
                                                     'email': [email_address],
                                                     'name': full_name
                                                     }

        return self.store_dict

    def append_and_return(self, sheet_obj, row, col):
        val = sheet_obj.cell(row=row, column=col).value
        self.row_vals.append(val)
        return val

    def check_validity(self):
        logging.basicConfig(filename='./logs/conflicts.log', encoding='utf-8', level=logging.DEBUG)
        for company in self.store_dict.keys():
            sub_list = [i for i in self.store_dict[company]['email'] if not isinstance(i, int)]
            if len(sub_list) > 1:
                logging.info(f'Multiple emails for company: {company}. Emails found: {sub_list}')


if __name__ == "__main__":
    ss = GetData(initials='JR')
    dr = ss.get_data()
    ss.check_validity()
