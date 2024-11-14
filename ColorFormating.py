from openpyxl.styles import Font
from openpyxl import Workbook
import datetime
import openpyxl

# Open the workbook and select the first sheet

def color_formatting(wb, ws, filename):
    """
    format excel cell based on rules like.

    Rule 1: Update Alert column text color to RED if created date is more than 45 days.
    Rule 2: Update D,E,F columns colours to Yellow if cell is blank.
    Rule 3: Update E column color to Yellow if cell containc XXX in value.
    :return:
    updated file name to share with business.
    """

    red_ft = Font(color="00FF0000")
    yellow_ft = openpyxl.styles.colors.Color(rgb='00FFF000')
    yellow_fill = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=yellow_ft)
    for row in ws.iter_rows(min_row=2):
        try:
            alert_cell = row[0]
            date_cell = row[1]
            sev_cell = row[3]
            dept_cell = row[4]
            issue_cell = row[5]

            delta = datetime.datetime.today() - date_cell.value
            if int(delta.days) > 45:
                alert_cell.font = red_ft
            if sev_cell.value is None:
                sev_cell.fill = yellow_fill
            if dept_cell.value is None or 'XXX' in str(dept_cell.value).upper():
                dept_cell.fill = yellow_fill
            if issue_cell.value is None:
                issue_cell.fill = yellow_fill
            wb.save(filename)

        except Exception as e:
            print(e)