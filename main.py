# python 3.9
import pandas as pd
import tkinter as tk
import os

from tkinter.filedialog import askopenfilenames
from tkinter import messagebox

root = tk.Tk()
root.title('엑셀파일통합')
root.minsize(400, 300)  # 최소 사이즈

dir_path = None  # 폴더 경로 담을 변수 생성
file_list = []  # 파일 목록 담을 변수 생성
writer = None

def integrationFiles():
    df = pd.DataFrame()
    for index in range(0, len(file_list)):
        file = file_list[index]
        if file.endswith('.csv'):
            f = pd.read_csv(file, skiprows=84, header=0, usecols=[1, 2], engine='python')
        else:
            f = pd.read_csv(file, skiprows=84, header=0, usecols=[1, 2], engine='openpyxl')
        df = pd.concat([df, f], axis=1, ignore_index=True)

    df.to_csv('./result.csv')
    messagebox.showwarning("완료", "파일 통합이 완료되었습니다.")


def integrationFiles2():
    df = pd.DataFrame()
    for index in range(0, len(file_list)):
        file = file_list[index]
        if file.endswith('.csv'):
            f = pd.read_csv(file, header=0, usecols=[4, 5], engine='python', sheet_name=4)
            print(f)
        elif file.endswith('.xlsx'):
            f = pd.read_excel(file, header=0, usecols=[4, 5], engine='openpyxl', sheet_name=4)
            print(f)
        else:
            xls_file = pd.ExcelFile(file)
            sheet_names = xls_file.sheet_names

            # Create dict
            res = {}

            # Build dict of sheetname: dataframe of each sheet
            for sheet in sheet_names:
                res[sheet] = pd.read_excel(file, sheet_name=sheet, header=None)

            # Create ExcelWriter object
            writer = pd.ExcelWriter("./temp.xlsx", engine='xlsxwriter')

            # Loop through dict, and have the writer write them to a single file
            for sheet, frame in res.items():
                frame.to_excel(writer, sheet_name=sheet, header=None, index=None)

            # Save off
            writer.save()
            f = pd.read_excel("./temp.xlsx", header=0, usecols=[4, 5], engine='openpyxl', sheet_name=4)

        df = pd.concat([df, f], axis=1, ignore_index=True)

    writer.close()
    os.remove("./temp.xlsx")
    df.to_csv('./result.csv')
    messagebox.showwarning("완료", "파일 통합이 완료되었습니다.")


def selectFiles():
    try:
        fileNames = askopenfilenames(initialdir="./", filetypes=(("Excel files", ".xlsx .xls"), ("Csv files", ".csv"), ('All files', '*.*')))

        if fileNames == '':
            messagebox.showwarning("경고", "파일을 선택 하세요")
        else:
            if len(fileNames) == 0:
                messagebox.showwarning("경고", "파일이 없습니다.")
            else:
                for fileName in fileNames:
                    file_list.append(fileName)
                    multipleFileListBox.insert(0, fileName)
    except:
        messagebox.showerror("Error", "오류가 발생했습니다.")
        multipleFileListBox.delete(0, "end")


def reset():
    try:
        reply = messagebox.askyesno("초기화", "정말로 초기화 하시겠습니까?")
        if reply:
            multipleFileListBox.delete(0, "end")
            messagebox.showinfo("Success", "초기화 되었습니다.")
    except:
        messagebox.showerror("Error", "오류가 발생했습니다.")


# 상단 프레임 (LabelFrame)
topFrame = tk.LabelFrame(root, pady=15, padx=15)
topFrame.grid(row=0, column=0, pady=10, padx=10, sticky="nswe")
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

# 하단 프레임 (Frame)
bottomFrame = tk.Frame(root, pady=10)
bottomFrame.grid(row=1, column=0, pady=10)

# 레이블
multipleFileLabel = tk.Label(topFrame, text='파일 선택')
# columnLabel = tk.Label(topFrame, text='열 지정')
# rowLable = tk.Label(topFrame, text='시작할 행 지정')

# 리스트박스
multipleFileListBox = tk.Listbox(topFrame, width=40)
# columnListBox = tk.Listbox(topFrame, width=40)
# rowListBox = tk.Listbox(topFrame, width=40)

# 버튼
multipleFileBtn = tk.Button(topFrame, text="추가", width=8, command=selectFiles)
resetBtn = tk.Button(bottomFrame, text="초기화", width=8, command=reset)
integrationBtn = tk.Button(bottomFrame, text="통합", width=8, command=integrationFiles2)

# 스크롤바 - 기능 연결
scrollbar = tk.Scrollbar(topFrame)
scrollbar.config(command=multipleFileListBox.yview)
multipleFileListBox.config(yscrollcommand=scrollbar.set)

# 상단 프레임
multipleFileLabel.grid(row=2, column=0, sticky="n")
multipleFileListBox.grid(row=2, column=1, rowspan=2, sticky="wens")

# columnLabel.grid(row=2, column=0, sticky="n")
# columnListBox.grid(row=2, column=1, rowspan=2, sticky="wens")

# rowLable.grid(row=2, column=0, sticky="n")
# rowListBox.grid(row=2, column=1, rowspan=2, sticky="wens")

scrollbar.grid(row=2, column=2, rowspan=2, sticky="wens")
multipleFileBtn.grid(row=2, column=3, sticky="n")

# 상단프레임 grid (2,1)은 창 크기에 맞춰 늘어나도록
topFrame.rowconfigure(2, weight=1)
topFrame.columnconfigure(1, weight=1)

# 하단 프레임
resetBtn.pack(side="left")
integrationBtn.pack(side="right")

root.mainloop()
