from __future__ import print_function
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError



class ggs:
    def __init__(self, creds) :
        self.creds = creds
        print(f"ggs class signed with cred : {self.creds.service_account_email}")
        

    def create(self, title):
        """
        Creates the Sheet the user has access to.
        Load pre-authorized user credentials from the environment.
        TODO(developer) - See https://developers.google.com/identity
        for guides on implementing OAuth2 for the application.
            """
        
        # pylint: disable=maybe-no-member
        try:
            service = build('sheets', 'v4', credentials=self.creds)
            spreadsheet = {
                'properties': {
                    'title': title
                }
            }
            spreadsheet = service.spreadsheets().create(body=spreadsheet,
                                                        fields='spreadsheetId') \
                .execute()
            print(f"Spreadsheet ID: {(spreadsheet.get('spreadsheetId'))}")
            return spreadsheet.get('spreadsheetId')
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error
        
    def get_values(self, spreadsheet_id, range_name):
        """
        Creates the batch_update the user has access to.
        Load pre-authorized user credentials from the environment.
        TODO(developer) - See https://developers.google.com/identity
        for guides on implementing OAuth2 for the application.
            """
        
        try:
            service = build('sheets', 'v4', credentials=self.creds)

            result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id, range=range_name).execute()
            rows = result.get('values', [])
            print(f"{len(rows)} rows retrieved")
            return result
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error
        

    def batch_get_values(self, spreadsheet_id, _range_names):
        """
        Creates the batch_update the user has access to.
        Load pre-authorized user credentials from the environment.
        TODO(developer) - See https://developers.google.com/identity
        for guides on implementing OAuth2 for the application.
            """
        
        # pylint: disable=maybe-no-member
        try:
            service = build('sheets', 'v4', credentials=self.creds)
            range_names = [
                # Range names ...
            ]
            result = service.spreadsheets().values().batchGet(
                spreadsheetId=spreadsheet_id, ranges=range_names).execute()
            ranges = result.get('valueRanges', [])
            print(f"{len(ranges)} ranges retrieved")
            return result
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error
        

    def update_values(self, spreadsheet_id, range_name, value_input_option,
                    _values):
        """
        Creates the batch_update the user has access to.
        Load pre-authorized user credentials from the environment.
        TODO(developer) - See https://developers.google.com/identity
        for guides on implementing OAuth2 for the application.
            """
        
        # pylint: disable=maybe-no-member
        try:

            service = build('sheets', 'v4', credentials=self.creds)
            values = [
                [
                    # Cell values ...
                ],
                # Additional rows ...
            ]
            body = {
                'values': values
            }
            result = service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id, range=range_name,
                valueInputOption=value_input_option, body=body).execute()
            print(f"{result.get('updatedCells')} cells updated.")
            return result
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error
        
    def batch_update_values(self, spreadsheet_id, range_name,
                            value_input_option, _values):
        """
            Creates the batch_update the user has access to.
            Load pre-authorized user credentials from the environment.
            TODO(developer) - See https://developers.google.com/identity
            for guides on implementing OAuth2 for the application.
                """
        
        # pylint: disable=maybe-no-member
        try:
            service = build('sheets', 'v4', credentials=self.creds)

            values = [
                [
                    # Cell values ...
                ],
                # Additional rows
            ]
            data = [
                {
                    'range': range_name,
                    'values': values
                },
                # Additional ranges to update ...
            ]
            body = {
                'valueInputOption': value_input_option,
                'data': data
            }
            result = service.spreadsheets().values().batchUpdate(
                spreadsheetId=spreadsheet_id, body=body).execute()
            print(f"{(result.get('totalUpdatedCells'))} cells updated.")
            return result
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error

    def append_values(self, spreadsheet_id, range_name, value_input_option,
                    _values):
        """
        Creates the batch_update the user has access to.
        Load pre-authorized user credentials from the environment.
        TODO(developer) - See https://developers.google.com/identity
        for guides on implementing OAuth2 for the application.
            """
        
        # pylint: disable=maybe-no-member
        try:
            service = build('sheets', 'v4', credentials=self.creds)

            values = [
                [
                    # Cell values ...
                ],
                # Additional rows ...
            ]
            body = {
                'values': values
            }
            result = service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id, range=range_name,
                valueInputOption=value_input_option, body=body).execute()
            print(f"{(result.get('updates').get('updatedCells'))} cells appended.")
            return result

        except HttpError as error:
            print(f"An error occurred: {error}")
            return error
        
    def sheets_batch_update(self, spreadsheet_id, title, find, replacement):

        """
        Update the sheet details in batch, the user has access to.
        Load pre-authorized user credentials from the environment.
        TODO(developer) - See https://developers.google.com/identity
        for guides on implementing OAuth2 for the application.
        """

        # pylint: disable=maybe-no-member

        try:
            service = build('sheets', 'v4', credentials=self.creds)

            requests = []
            # Change the spreadsheet's title.
            requests.append({
                'updateSpreadsheetProperties': {
                    'properties': {
                        'title': title
                    },
                    'fields': 'title'
                }
            })
            # Find and replace text
            requests.append({
                'findReplace': {
                    'find': find,
                    'replacement': replacement,
                    'allSheets': True
                }
            })
            # Add additional requests (operations) ...

            body = {
                'requests': requests
            }
            response = service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body).execute()
            find_replace_response = response.get('replies')[1].get('findReplace')
            print('{0} replacements made.'.format(
                find_replace_response.get('occurrencesChanged')))
            return response

        except HttpError as error:
            print(f"An error occurred: {error}")
            return error

    def pivot_tables(self, spreadsheet_id):
        """
        Creates the batch_update the user has access to.
        Load pre-authorized user credentials from the environment.
        TODO(developer) - See https://developers.google.com/identity
        for guides on implementing OAuth2 for the application.
            """
        # pylint: disable=maybe-no-member
        try:
            service = build('sheets', 'v4', credentials=self.creds)
            # Create two sheets for our pivot table.
            body = {
                'requests': [{
                    'addSheet': {}
                }, {
                    'addSheet': {}
                }]
            }
            batch_update_response = service.spreadsheets() \
                .batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
            source_sheet_id = batch_update_response.get('replies')[0] \
                .get('addSheet').get('properties').get('sheetId')
            target_sheet_id = batch_update_response.get('replies')[1] \
                .get('addSheet').get('properties').get('sheetId')
            requests = []
            requests.append({
                'updateCells': {
                    'rows': {
                        'values': [
                            {
                                'pivotTable': {
                                    'source': {
                                        'sheetId': source_sheet_id,
                                        'startRowIndex': 0,
                                        'startColumnIndex': 0,
                                        'endRowIndex': 20,
                                        'endColumnIndex': 7
                                    },
                                    'rows': [
                                        {
                                            'sourceColumnOffset': 1,
                                            'showTotals': True,
                                            'sortOrder': 'ASCENDING',

                                        },

                                    ],
                                    'columns': [
                                        {
                                            'sourceColumnOffset': 4,
                                            'sortOrder': 'ASCENDING',
                                            'showTotals': True,

                                        }
                                    ],
                                    'values': [
                                        {
                                            'summarizeFunction': 'COUNTA',
                                            'sourceColumnOffset': 4
                                        }
                                    ],
                                    'valueLayout': 'HORIZONTAL'
                                }
                            }
                        ]
                    },
                    'start': {
                        'sheetId': target_sheet_id,
                        'rowIndex': 0,
                        'columnIndex': 0
                    },
                    'fields': 'pivotTable'
                }
            })
            body = {
                'requests': requests
            }
            response = service.spreadsheets() \
                .batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
            return response

        except HttpError as error:
            print(f"An error occurred: {error}")
            return error
        

    def conditional_formatting(self, spreadsheet_id):
        """
        Creates the batch_update the user has access to.
        Load pre-authorized user credentials from the environment.
        TODO(developer) - See https://developers.google.com/identity
        for guides on implementing OAuth2 for the application.
            """
        # pylint: disable=maybe-no-member
        try:
            service = build('sheets', 'v4', credentials=self.creds)

            my_range = {
                'sheetId': 0,
                'startRowIndex': 1,
                'endRowIndex': 11,
                'startColumnIndex': 0,
                'endColumnIndex': 4,
            }
            requests = [{
                'addConditionalFormatRule': {
                    'rule': {
                        'ranges': [my_range],
                        'booleanRule': {
                            'condition': {
                                'type': 'CUSTOM_FORMULA',
                                'values': [{
                                    'userEnteredValue':
                                        '=GT($D2,median($D$2:$D$11))'
                                }]
                            },
                            'format': {
                                'textFormat': {
                                    'foregroundColor': {'red': 0.8}
                                }
                            }
                        }
                    },
                    'index': 0
                }
            }, {
                'addConditionalFormatRule': {
                    'rule': {
                        'ranges': [my_range],
                        'booleanRule': {
                            'condition': {
                                'type': 'CUSTOM_FORMULA',
                                'values': [{
                                    'userEnteredValue':
                                        '=LT($D2,median($D$2:$D$11))'
                                }]
                            },
                            'format': {
                                'backgroundColor': {
                                    'red': 1,
                                    'green': 0.4,
                                    'blue': 0.4
                                }
                            }
                        }
                    },
                    'index': 0
                }
            }]
            body = {
                'requests': requests
            }
            response = service.spreadsheets() \
                .batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
            print(f"{(len(response.get('replies')))} cells updated.")
            return response

        except HttpError as error:
            print(f"An error occurred: {error}")
            return error

    def filter_views(self, spreadsheet_id):
        """
            Creates the batch_update the user has access to.
            Load pre-authorized user credentials from the environment.
            TODO(developer) - See https://developers.google.com/identity
            for guides on implementing OAuth2 for the application.
                """
        
        # pylint: disable=maybe-no-member
        try:
            service = build('sheets', 'v4', credentials=self.creds)

            my_range = {
                'sheetId': 0,
                'startRowIndex': 0,
                'startColumnIndex': 0,
            }
            addfilterviewrequest = {
                'addFilterView': {
                    'filter': {
                        'title': 'Sample Filter',
                        'range': my_range,
                        'sortSpecs': [{
                            'dimensionIndex': 3,
                            'sortOrder': 'DESCENDING'
                        }],
                        'criteria': {
                            0: {
                                'hiddenValues': ['Panel']
                            },
                            6: {
                                'condition': {
                                    'type': 'DATE_BEFORE',
                                    'values': {
                                        'userEnteredValue': '4/30/2016'
                                    }
                                }
                            }
                        }
                    }
                }
            }

            body = {'requests': [addfilterviewrequest]}
            addfilterviewresponse = service.spreadsheets() \
                .batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

            duplicatefilterviewrequest = {
                'duplicateFilterView': {
                    'filterId':
                        addfilterviewresponse['replies'][0]
                        ['addFilterView']['filter']
                        ['filterViewId']
                }
            }

            body = {'requests': [duplicatefilterviewrequest]}
            duplicatefilterviewresponse = service.spreadsheets() \
                .batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

            updatefilterviewrequest = {
                'updateFilterView': {
                    'filter': {
                        'filterViewId': duplicatefilterviewresponse['replies'][0]
                        ['duplicateFilterView']['filter']['filterViewId'],
                        'title': 'Updated Filter',
                        'criteria': {
                            0: {},
                            3: {
                                'condition': {
                                    'type': 'NUMBER_GREATER',
                                    'values': {
                                        'userEnteredValue': '5'
                                    }
                                }
                            }
                        }
                    },
                    'fields': {
                        'paths': ['criteria', 'title']
                    }
                }
            }

            body = {'requests': [updatefilterviewrequest]}
            updatefilterviewresponse = service.spreadsheets() \
                .batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
            print(str(updatefilterviewresponse))
        except HttpError as error:
            print(f"An error occurred: {error}")


