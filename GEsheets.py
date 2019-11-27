import requests
import pandas as pd
import numpy as np
import json


class SS:  # access sheets

    def __init__(self, apikey, proxy=True, project_code=""):  # When using the class, input apikey, proxy is set to true by default, if not connected to GE network - select false. 
    														  # Project codes are also optional, they don't really do anything unless you call a list - if there is a project code 
    														  # it only shows sheets containing the code in the title.
        self.project_code = project_code

        global url               # Set as global variables because I don't know how to call variables in a class that's already in a class.
        global headers
        global parameters
        global proxies

        url = "https://api.smartsheet.com/2.0/sheets/"

        headers = {
            'Authorization': "Bearer " + str(apikey), 
            'Content-Type': "application/json"
        }

        parameters = {
            'includeAll': "True",
            'loadAll': "True"
        }

        if proxy == True:
            proxies = {'https': 'https://PITC-Zscaler-EMEA-Amsterdam3PR.proxy.corporate.ge.com:80'}
        else:
            proxies = None

        response = requests.get(url, headers=headers, proxies=proxies, params=parameters)

        data = response.json()
        df = pd.DataFrame(data['data'], columns=['name', 'id', 'modifiedAt', 'permalink'])      

        self.data = data
        self.list_sheets = df
        self.btp = url

    class sheet:  # This class looks inside sheets

        def __init__(self, sheet_id):  # The class is initialized with the sheet ID, all the below functions only work inside a sheet

            self.sheet_id = sheet_id

        def view_sheet(self):  # Complete view of sheet and contains all data

            sheet_data = requests.get(url + str(self.sheet_id), headers=headers, proxies=proxies).json()

            if (list(sheet_data.keys())[0] == 'errorCode') is True:  # If the code is incorrect, return an error
                sheet_data = 'Please check the sheet ID'
            else:
                pass

            self.sheet_data = sheet_data

            return sheet_data

        def columns(self):  # creates a list of column names with respective IDs, in later functions columns are edited with ID's instead of names, this can be changed if needs be

            response = requests.get(url + str(self.sheet_id) + '/columns', headers=headers, proxies=proxies).json()
            df_cols = pd.DataFrame(response['data'], columns=['title', 'id'])

            return df_cols

        def view_column(self, column_id): # Describes the properties of the column

            response = requests.get(url + str(self.sheet_id) + '/columns/' + str(column_id), headers=headers,
                                    proxies=proxies).json()

            return response

        def getColumnId(self, column_name):

        	return int(self.columns()[self.columns()['title'] == column_name]['id'])

        def payload(self, row_id="", column_id="", value="", formula=""):  # This is used to create input data in the correct JSON format which smartsheet can read

            payload = {"id": row_id, "cells": [{"columnId": str(column_id), "value": str(value), "formula": formula}]}

            return payload

        def put_cell(self, payload):  # Used with the payload function above, this updates the cell with the new data

            response = requests.put(url + str(self.sheet_id) + '/rows/', headers=headers, proxies=proxies,
                                    params=parameters, data=json.dumps(payload))

            return response

        def rows(self):  # Creates an array of every row ID in the sheet

            try:
                row_ids = []
                for i in self.sheet_data['rows']:
                    row_ids.append(i['id'])
            except:

                row_ids = []
                for i in self.view_sheet()['rows']:
                    row_ids.append(i['id'])

            return row_ids

        def view_rows(self): # Displays the data inside each row

            try:
                return self.sheet_data['rows']
            except:
                return self.view_sheet()['rows']

        def delete_empty(self, column_title):  # Deletes any instances of rows with blank values in the specified column

            
            sheet = self.view_sheet()
            column = self.columns()[self.columns()['title'] == column_title]['id'].values[0]

            empty_cells = []
            for i in sheet['rows']:   # Goes through every cell in every row and checks if it's blank, if it is, append the row ID to the list to delete.

                for j in i['cells']:

                    if j['columnId'] == column:
                        try:
                            check_blank = j['value']

                        except:
                            empty_cells.append(i['id'])

            seperator = ','
            empty_cells_str = seperator.join(str(value) for value in empty_cells)

            response = requests.delete(
                'https://api.smartsheet.com/2.0/sheets/' + str(self.sheet_id) + '/rows/?ids=' + empty_cells_str,
                headers=headers, proxies=proxies).json()

        def convert_formula(self):  # only set up for one type of column - do not use for others - changes the formula in the status indicator columnhbv
        	

	        try:

	            
	            rows = self.rows()

	            column = self.columns()[self.columns()['title'] == 'Status Indicator']['id'].values[0]

	            unlock = {"locked": "false", "type": "TEXT_NUMBER"}
	            requests.put('https://api.smartsheet.com/2.0/sheets/' + str(self.sheet_id) + '/columns/' + str(column),
	                         headers=headers, proxies=proxies, params=parameters, data=json.dumps(unlock))
	            
	            payloads = []

	            for i in range(0, len(rows)):
	                row_no = i + 1

	                formula = ("""=IF(AND([Planned Finish Date]" + str(row_no) + " > [Baseline / Due date]" + str(row_no) + ",[Baseline / Due date]" + str(row_no) + "<> ""\"\"""), ""\"Red\""", 
	                			IF([% Complete]" + str(row_no) + " = 0, ""\"\""", 
	                			IF(AND([% Complete]" + str(row_no) + " > 0, [% Complete]" + str(row_no) + " < 1), ""\"Yellow\""", 
	                			IF([% Complete]" + str(row_no) + " = 1, ""\"Green\""", ""\"\"""))))"""
	                # formula = ("Red")

	                payloads.append(self.payload(rows[i], column, formula=formula))

	            self.put_cell(payloads)

	            lock = {"locked": "true", "type": "PICKLIST", "symbol": "RYG"}
	            # requests.put('https://api.smartsheet.com/2.0/sheets/' + str(self.sheet_id) + '/columns/' + str(column),
	            #              headers=headers, proxies=proxies, params=parameters, data=json.dumps(lock))

	            print('Completed sheet: ' + str(self.sheet_id) + 'g')
	        except:
	        	print('Error')

	         	





