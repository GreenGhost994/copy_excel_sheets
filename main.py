from win32com.client import Dispatch

source = r"C:\Users\piotr.zielinski\Downloads\Global design schedule residential (broken).xlsm"
output = r"C:\Users\piotr.zielinski\Downloads\Global design schedule residential (fixed3).xlsm"

def copy_sheets(source_filename: str, output_filename: str):
    xl=Dispatch("Excel.Application")
    xl.DisplayAlerts=False
    xl.EnableEvents = False
    xl.Visible=True  # You can remove this line if you don't want the Excel application to be visible

    source_wb=xl.Workbooks.Open(Filename=source_filename)
    output_wb=xl.Workbooks.Open(Filename=output_filename)

    source_sheet_names=[sheet.Name for sheet in source_wb.Sheets]
    output_sheet_names=[sheet.Name for sheet in output_wb.Sheets]

    not_copy = ['SCHEDULE', 'REALIZED PRODUCTION', 'CHART EXTRA DATA', 'PROJECT_TEMP', 'LISTS', 'SCHEDULE_TEMP']
    i = 0
    for sheet_name in enumerate(output_sheet_names):
        if i < 500:
            i += 1
            if sheet_name[1] in source_sheet_names:
                if sheet_name[1] not in not_copy:
                    try:
                        lastColumn=output_wb.Worksheets(sheet_name[1]).UsedRange.Columns.Count
                        lastRow=output_wb.Worksheets(sheet_name[1]).UsedRange.Rows.Count

                        s_ws = source_wb.Worksheets(sheet_name[1])
                        o_ws = output_wb.Worksheets(sheet_name[1])

                        s_ws.Range(s_ws.Cells(1, 1), s_ws.Cells(lastRow, lastColumn)).Copy(o_ws.Range(o_ws.Cells(1, 1),
                                     o_ws.Cells(lastRow, lastColumn)))
                    except Exception as e:
                        print('Problem', sheet_name[1], e)
                        pass
            else:
                print(f'Missing {sheet_name[1]}')
                pass

    output_wb.Close(SaveChanges=True)
    xl.Quit()


if __name__ == "__main__":
   copy_sheets(source, output)
