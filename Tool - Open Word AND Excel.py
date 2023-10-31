import os
from win32com.client import Dispatch
file_name = "Test"
file_content1 = ""
file_directory = "D:\\ALL\\Password\\8\\8\\8\\8\\8\\8\\Hello-Lawrence\\Other\\pythonProject"
excel_coordinates = "A1"


def word():
    word = Dispatch('Word.Application')
    # 或者使用下面的方法，使用启动独立的进程：
    # word = DispatchEx('Word.Application')

    # 如果不声明以下属性，运行的时候会显示的打开word
    word.Visible = 1  # 0:后台运行 1:前台运行(可见)
    word.DisplayAlerts = 0  # 不显示，不警告

    # 创建新的word文档
    doc = word.Documents.Add()

    bottom = doc.Range()
    bottom.InsertAfter(file_content1)

    doc.SaveAs(file_directory + "\\" + file_name + ".docx")  # 另存为
    doc.Close()  # 关闭 word 文档
    word.Quit()  # 关闭 office


def excel():
    excel = Dispatch('Excel.Application')  # 获取Excel
    excel.Visible = 1
    wb = excel.Workbooks.Add()
    ws = wb.Worksheets('Sheet1')

    # c_column = [20, 30, 20, 30, 40, 50, 35, 45]
    #
    # d_column = [15, 31, 22, 15, 16, 40, 34, 89]
    #
    # ws.Range('C2').value = c_column[0]
    # ws.Range('E2').value = d_column[0]

    def R(alphabet, number, value):
        alphabet.capitalize()
        ws.Range(alphabet + number).value = value

    r1 = input("Enter range 1  | ")
    r2 = input("Enter range 2  | ")
    val = input("Enter value    | ")
    R(r1, r2, value=val)
    wb.SaveAs(file_directory + "\\" + file_name + ".xlsx")
    wb.Close()
    excel.Quit()


def noCorrectInput():
    print("Your input needs to be 1 or 2. Please do it again. ")
    os.system('cls')


if __name__ == '__main__':
    Type = int(input("Word(1) or Excel(2)?"))
    if Type == 1:
        file_name = input("File name      | ")
        file_content1 = input("File Line 1    | ")
        word()
    elif Type == 2:
        file_name = input("File name      | ")
        excel()
    else:
        noCorrectInput()
