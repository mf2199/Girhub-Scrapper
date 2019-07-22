import github_utils
import sheet
import datetime
import time


HOUR_DURATION = 216000


# sheet_id = sheet.create_new_sheet(
#     'QLogic Python Internal Review'
# )
# sheet.create_title_row(
#     sheet_id,
#     [
#         ['Priority', 80, 'CENTER', ['Paused', 'Low', 'Medium', 'High', 'Critical', 'Done']],
#         ['Issue', 50, 'CENTER', None],
#         ['Status', None, 'CENTER', ['PENDING', 'OPEN', 'SUBMITTED', 'ON HOLD', 'MERGED', 'REOPENED', 'CLOSED', 'IRRELEVANT']],
#         ['Created', None, 'CENTER', None],
#         ['Description', 450, None, None],
#         ['Repository', None, 'CENTER', None],
#         ['API', None, 'CENTER', None],
#         ['Assignee', None, 'CENTER', ['Ilya', 'Hemang', 'Maxim', 'Paras', 'Sumit', 'Sangram', 'N/A']],
#         ['Task', None, 'CENTER', None],
#         ['Opened', None, 'CENTER', None],
#         ['Internal PR', None, 'CENTER', None],
#         ['Public PR', None, 'CENTER', None],
#         ['Comment', 550, None, None],
#     ]
# )
# rows, count = github_utils.build_whole_table()
# sheet.save_to_sheet(sheet_id, rows, count)
# print(sheet.read_sheet(sheet_id, 'Sheet1'))

# updating table every hour
sheet_id = '1_ENbCfw7V5kY32jLM3hUlwXxv80poIHpXFdfysM9r-s'

while True:
    print('update: ' + str(datetime.datetime.now()))
    sheet.update_list(sheet_id, 'Sheet1')
    print('update successful')

    time.sleep(HOUR_DURATION)
