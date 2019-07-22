import string
import github_utils
import operator


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


class Columns:
    def __init__(self, cols_list):
        self.names = []
        self.size_requests = []
        self.align_requests = []
        self.one_of_requests = []

        for index, col in enumerate(cols_list):
            self.names.append(col[0])

            if col[1] is not None:
                self.size_requests.append(
                    self._gen_size_request(index, col[1])
                )

            if col[2] in ('CENTER', 'LEFT', 'RIGHT'):
                self.align_requests.append(
                    self._gen_align_request(index, col[2])
                )

            if col[3] is not None:
                self.one_of_requests.append(
                    self._gen_one_of_request(index, col[3])
                )

    @property
    def sym_range(self):
        sym_range = 'A1:{last_letter}1'.format(
            last_letter=string.ascii_uppercase[
                len(self.names) - 1
            ]
        )
        return sym_range

    @property
    def total_row_request(self):
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
    return url.split(';')[1][1:-2]

def _build_url(num, repo_lts):
    repo = REPO_NAMES[repo_lts]
    url = '=ГИПЕРССЫЛКА("https://github.com/{repo}/issues/{num}";"{num}")'.format(
        repo=repo, num=num
    )
    return url


def create_new_sheet(service, title):
    """Create new Google Sheet with given title."""
    body = {'properties': {'title': title}}

    spreadsheet = service.spreadsheets().create(
        body=body,
        fields='spreadsheetId'
    ).execute()

    return spreadsheet.get('spreadsheetId')


def create_title_row(service, sheet_id, cols_list):
    requests = []

    columns = Columns(cols_list)
    body = {'values': [columns.names]}

    service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=columns.sym_range,
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()

    requests += columns.total_row_request
    requests += columns.size_requests
    requests += columns.align_requests
    requests += columns.one_of_requests

    body = {'requests': requests}

    service.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body=body
    ).execute()


def read_sheet(service, sheet_id, range_):
    result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id, range=range_
    ).execute()
    return result['values']


def update_list(service, sheet_id, list_name):
    table = read_sheet(service, sheet_id, list_name)[1:]
    table.sort(key=operator.itemgetter(5, 6, 1))
    in_table = [(row[1], row[5]) for row in table]

    new_data, count = github_utils.build_whole_table()
    in_new = [(_get_num_from_url(row[1]), row[5]) for row in new_data]

    num = 0
    requests = []

    while num < max(len(table), len(new_data)):
        try:
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

        new_num = _get_num_from_url(new_data[num][1])
        if not (new_num, new_data[num][5]) in in_table:
            table.insert(num, new_data[num])

        for field in github_utils.TRACKED_FIELDS:
            table[num][field] = new_data[num][field]

        num += 1

    table.sort(key=sort_func)
    save_to_sheet(service, sheet_id, table, count)

    if requests:
        body = {'requests': requests}

        service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body=body
        ).execute()


def sort_func(item):
    return item[5], item[6], int(_get_num_from_url(item[1]))


def save_to_sheet(service, sheet_id, rows, count):
    body = {'values': rows}

    result = service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range="A2:H{count}".format(count=count),
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()
