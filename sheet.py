"""
Funtions and objects, which uses Google Sheets API to
build review tables.
"""
import string
import github_utils
import operator
import auth


REPO_NAMES = {
    'GCP': 'googleapis/google-cloud-python',
    'GRMP': 'googleapis/google-resumable-media-python'
}


REPO_NAMES_INV = {
    'googleapis/google-cloud-python': 'GCP',
    'googleapis/google-resumable-media-python': 'GRMP'
}

RED = {
    'red': 1,
    'green': 0.38,
    'blue': 0.52,
    'alpha': 1
}


service = auth.authenticate()


class Columns:
    """Object for column processing and generating requests."""
    def __init__(self, cols_list):
        self.names = []
        self._requests = []

        # generating requests from user's settings
        for index, col in enumerate(cols_list):
            self.names.append(col[0])

            if col[1] is not None:
                self._requests.append(
                    self._gen_size_request(index, col[1])
                )

            if col[2] in ('CENTER', 'LEFT', 'RIGHT'):
                self._requests.append(
                    self._gen_align_request(index, col[2])
                )

            if col[3] is not None:
                self._requests.append(
                    self._gen_one_of_request(index, col[3])
                )

    @property
    def requests(self):
        self._requests.append(self._title_row_request)
        return self._requests

    @property
    def sym_range(self):
        """Return sumbolic range pointers for columns."""
        sym_range = 'A1:{last_letter}1'.format(
            last_letter=string.ascii_uppercase[
                len(self.names) - 1
            ]
        )
        return sym_range

    @property
    def _title_row_request(self):
        """Bolding and aligning title row."""
        request = [{
            'repeatCell': {
                'fields': 'userEnteredFormat',
                'range': {
                    'sheetId': 0,
                    'startRowIndex': 0,
                    'endRowIndex': 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': len(self.names)
                },
                'cell': {
                    'userEnteredFormat': {
                        'horizontalAlignment': "CENTER",
                        'textFormat': {'bold': True}
                    }
                }
            }
        }]
        return request

    def _gen_align_request(self, index, align):
        """Align set request for user-specified columns."""
        request = {
            'repeatCell': {
                'fields': 'userEnteredFormat',
                'range': {
                    'sheetId': 0,
                    'startRowIndex': 1,
                    'endRowIndex': 1000,
                    'startColumnIndex': index,
                    'endColumnIndex': index + 1
                },
                'cell': {
                    'userEnteredFormat': {
                        'horizontalAlignment': align
                    }
                }
            }
        }

        return request


    def _gen_size_request(self, index, size):
        """Column width request for user-specified columns."""
        request = {
            "updateDimensionProperties": {
                "properties": {"pixelSize": size},
                "fields": "pixelSize",
                "range": {
                    "sheetId": 0,
                    "dimension": "COLUMNS",
                    "startIndex": index,
                    "endIndex": index + 1,
                }
            }
        }

        return request

    def _gen_one_of_request(self, index, values):
        """Request to set data validation for user-specified columns."""
        vals = [{'userEnteredValue': value} for value in values]

        request = {
            "setDataValidation": {
                "range": {
                    "sheetId": 0,
                    'startRowIndex': 1,
                    'endRowIndex': 1000,
                    'startColumnIndex': index,
                    'endColumnIndex': index + 1,
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": vals
                    },
                "showCustomUi": True,
                "strict": True
                }
            }
        }

        return request


def _make_color_request(row, column, color):
    """Request, that changes color of specified cell."""
    request = {
        'repeatCell': {
            'fields': 'userEnteredFormat',
            'range': {
                'sheetId': 0,
                'startRowIndex': row,
                'endRowIndex': row + 1,
                'startColumnIndex': column,
                'endColumnIndex': column + 1
            },
            'cell': {
                'userEnteredFormat': {
                    'backgroundColor': RED,
                    'horizontalAlignment': "CENTER",
                }
            }
        }
    }
    return request


def _get_num_from_url(url):
    """Get issue number from its URL."""
    return url.split(';')[1][1:-2]

def _build_url(num, repo_lts):
    """Build URL from issue number and repo name."""
    repo = REPO_NAMES[repo_lts]
    url = '=ГИПЕРССЫЛКА("https://github.com/{repo}/issues/{num}";"{num}")'.format(
        repo=repo, num=num
    )
    return url


def create_new_sheet(title):
    """Create new Google Sheet with given title."""
    body = {'properties': {'title': title}}

    spreadsheet = service.spreadsheets().create(
        body=body,
        fields='spreadsheetId'
    ).execute()

    return spreadsheet.get('spreadsheetId')


def create_title_row(sheet_id, cols_list):
    """Creating title row for table."""
    columns = Columns(cols_list)
    body = {'values': [columns.names]}

    # writing data
    service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=columns.sym_range,
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()

    if requests:
        body = {'requests': columns.requests}

        service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body=body
        ).execute()


def read_sheet(sheet_id, range_):
    """Reading the specified existing sheet."""
    table = service.spreadsheets().values().get(
        spreadsheetId=sheet_id, range=range_
    ).execute()['values'][1:]
    # sorting table for cases if someone added issues manually
    table.sort(key=operator.itemgetter(5, 6, 1))
    return table


def update_list(sheet_id, list_name):
    """Updating specified list of Google Sheet document."""
    num = 0
    requests = []

    # reading existing table
    table = read_sheet(sheet_id, list_name)
    # ids of already tracked issues
    in_table = [(row[1], row[5]) for row in table]

    # building new table from repositories
    new_data, count = github_utils.build_whole_table()
    in_new = [(_get_num_from_url(row[1]), row[5]) for row in new_data]

    while num < max(len(table), len(new_data)):
        try:
            # checking if issue closed
            if not (table[num][1], table[num][5]) in in_new:
                requests.append(_make_color_request(num + 1, 1, RED))
                table[num][1] = _build_url(
                    table[num][1],
                    table[num][5]
                )
                num += 1
                continue
        except IndexError:
            pass

        # inserting new issues
        new_num = _get_num_from_url(new_data[num][1])
        if not (new_num, new_data[num][5]) in in_table:
            table.insert(num, new_data[num])

        # updating tracked fields in existing data
        for field in github_utils.TRACKED_FIELDS:
            table[num][field] = new_data[num][field]

        num += 1

    save_to_sheet(sheet_id, table, count)

    if requests:
        body = {'requests': requests}

        service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body=body
        ).execute()


def sort_func(item):
    return item[5], item[6], int(_get_num_from_url(item[1]))


def save_to_sheet(sheet_id, rows, count):
    """Update existing sheet with new data."""
    rows.sort(key=sort_func)
    body = {'values': rows}

    result = service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range="A2:H{count}".format(count=count),
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()
