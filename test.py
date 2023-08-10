import pandas as pd
import xlsxwriter
import sys
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkinterdnd2 import TkinterDnD, DND_FILES
import subprocess
import os
from collections import OrderedDict
import math

target_folder = ''

def run_script(event=None):
	global df_result
	
	sheet_stock = pd.ExcelFile(file_entry0.get()).sheet_names
	sheet_sold = pd.ExcelFile(file_entry1.get()).sheet_names
	
	if not '재고현황' in sheet_stock:
		show_popup("오류","재고현황 파일 선택란에 다른 엑셀 파일이 선택되었습니다. 재고현황 엑셀 파일을 선택해주세요.", "return")
		
	if not '판매현황' in sheet_sold:
		show_popup("오류","판매현황 파일 선택란에 다른 엑셀 파일이 선택되었습니다. 판매현황 엑셀 파일을 선택해주세요.", "return")
		
	df_stock = pd.read_excel(file_entry0.get(), sheet_name='재고현황')
	df_sold = pd.read_excel(file_entry1.get(), sheet_name='판매현황')
	
	array_sold = ''
	
	array_stock_pre = df_stock.iloc[1:-2, [2,3,4]].values.tolist()
	print(array_stock_pre)
	array_stock = []
	
	for key, value, amount in array_stock_pre:
		if str(value) != "nan":
			array_stock.append([str(key) + " [" + str(value) + "]", amount])
		else:
			array_stock.append([key, amount])
	
	print(array_stock)
	
	array_sold = df_sold.iloc[1:-2, [0,1]].values.tolist()
	
	second_column = ''
	third_column = ''
	
	result_array = []
	dict_array = {}
	
	if len(array_sold) > len(array_stock):
		second_column = '재고수량'
		third_column = '판매수량'
	
		for key, value in array_stock:
			dict_array[key] = [value, None]
			
		
		for key, value in array_sold:
			if key in dict_array:
				dict_array[key][1] = value
			else:
				dict_array[key] = [None, value]
	else:
		second_column = '판매수량'
		third_column = '재고수량'
		
		for key, value in array_sold:
			dict_array[key] = [value, None]
		
		for key, value in array_stock:
			if key in dict_array:
				dict_array[key][1] = value
			else:
				dict_array[key] = [None, value]

	for key, values in dict_array.items():
		result_array.append([key] + values)

	result_array = sorted(result_array, key=lambda x: x[0])
	
	df_result = pd.DataFrame(result_array, columns=['품목명',second_column,third_column])
	df_result = df_result.fillna('')
	print(df_result)

	show_popup("거의 다 됐습니다","만들어질 엑셀 파일을 저장할 폴더를 선택해주세요.", "select_folder")

def close_popup(cmd):
	popup.destroy()
	if cmd == "return":
		return
	if cmd == "select_folder":
		select_folder()

def show_popup(title,msg,cmd):
	global popup
	popup = tk.Toplevel(root)
	popup.title(title)
	popup.resizable(False,False)
	
	popup_label = tk.Label(popup, text=msg)
	popup_label.pack()

	close_button = tk.Button(popup, text="확인", command=lambda: close_popup(cmd))
	close_button.pack()

def open_file_dialog(index):
	if index == 0:
		file_path = filedialog.askopenfilename(filetypes=[("재고현황 엑셀 파일", "*.xlsx")])
		if file_path:
			file_entry0.delete(0, tk.END)
			file_entry0.insert(0, file_path)
	elif index == 1:
		file_path = filedialog.askopenfilename(filetypes=[("판매현황 엑셀 파일", "*.xlsx")])
		if file_path:
			file_entry1.delete(0, tk.END)
			file_entry1.insert(0, file_path)

def select_folder():
	folder_path = filedialog.asksaveasfilename(initialfile="새로운 엑셀문서",defaultextension=".xlsx")
	if not folder_path:
		return
	target_folder = folder_path
	writer = pd.ExcelWriter(target_folder, engine='xlsxwriter')
	df_result.to_excel(writer, index=False, sheet_name='Sheet1')
	writer.close()

	subprocess.call('open "{}"'.format(target_folder), shell=True)

def on_drop(event, entry):
	file_path = event.data.replace("{", "").replace("}", "")  # 중괄호 제거
	entry.delete(0, tk.END)
	entry.insert(0, file_path)
	
root = TkinterDnD.Tk()
root.title("재고현황과 판매현황을 하나의 엑셀로 합칩니다.")

root.geometry("800x150")
root.resizable(True,False)

file_button0 = tk.Button(root, text="재고현황 파일 선택", command=lambda: open_file_dialog(0))
file_button0.pack()

file_entry0 = tk.Entry(root)
file_entry0.pack(fill=tk.X)

file_entry0.drop_target_register(DND_FILES)
file_entry0.dnd_bind('<<Drop>>', lambda event, entry=file_entry0: on_drop(event, entry))

file_button1 = tk.Button(root, text="판매현황 파일 선택", command=lambda: open_file_dialog(1))
file_button1.pack()

file_entry1 = tk.Entry(root)
file_entry1.pack(fill=tk.X)

file_entry1.drop_target_register(DND_FILES)
file_entry1.dnd_bind('<<Drop>>', lambda event, entry=file_entry1: on_drop(event, entry))


button = tk.Button(root, text="작업 시작", command=run_script)
button.pack()

root.mainloop()
