try:
    print("in try")
    xlApp = win32com.client.Dispatch("Excel.Application")
    print("opened app")
    xlWbk = xlApp.Workbooks.Open(r"C:\Users\liam.obrien\Desktop\CTR Report Generator\Downloads\CTR_18-022.xml")
    print("opened workbook")
    xlWbk.SaveAs(r"C:\Users\liam.obrien\Desktop\CTR Report Generator\Downloads\Output.xlsx", 51)
    print("saved workbook")

    xlWbk.Close(True)
    xlApp.Quit()

except Exception as e:
    print(e)

finally:
    xlWbk = None;
    xlApp = None
    del xlWbk;
    del xlApp