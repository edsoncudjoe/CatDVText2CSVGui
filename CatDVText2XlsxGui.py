import Tkinter as tk
import tkFileDialog
import tkMessageBox
import ttk
import sys
import xlsxwriter
import os

N = tk.N
S = tk.S
E = tk.E
W = tk.W
END = tk.END

class CatDV2CSV(tk.Frame):
	"""
	Simple program to convert a CatDV text file into a CSV file for excel.
	"""

	def __init__(self, parent):
		tk.Frame.__init__(self, parent)
		self.parent = parent

		self.mainf = tk.Frame(parent, bg='gray93')
		self.mainf.grid(sticky=N+S+E+W, padx=5, pady=5)

		self.mainf.rowconfigure(0, weight=4)
		self.mainf.columnconfigure(0, weight=4)	

		self.create_menubar()
		self.create_widgets()
		self.grid_widgets()

		self.file_options = options = {}
		options['filetypes'] = [('text files', '.txt'), ('xlsx files', '.xlsx')]
		self.collection = []


	def create_menubar(self):
		pass

	def create_widgets(self):
		self.cdv_text_label = ttk.Label(self.mainf, 
			text="Open CatDV text file.")
		self.cdv_text_btn = ttk.Button(self.mainf, text="open",
			command=self.load_catdv_data)
		self.save_label = ttk.Label(self.mainf, text="Save file as.")
		self.save_btn = ttk.Button(self.mainf, text="Save As")
		self.delete_text_label = ttk.Label(self.mainf, 
			text="Delete CatDV text file?")
		self.delete_text_btn = ttk.Button(self.mainf, text="delete text")
		self.quit_btn = ttk.Button(self.mainf, text="quit")

	def grid_widgets(self):
		self.cdv_text_label.grid(row=0, column=0)
		self.cdv_text_btn.grid(row=0, column=1)
		self.save_label.grid(row=1, column=0)
		self.save_btn.grid(row=1, column=1)
		self.delete_text_label.grid(row=2, column=0)
		self.delete_text_btn.grid(row=2, column=1)
		self.quit_btn.grid(row=3, column=1)

	def ask_open_file(self):
		"""Returns selected file in read mode"""
		self.cdv_file = tkFileDialog.askopenfile(mode='r', **self.file_options)
		return self.cdv_file

	def collect_titles(self):
		"""Collects text file information into a list"""
		for item in self.cdv_file:
			if item.startswith('IV'):
				item = item.split()
				self.collection.append(item)

	def load_catdv_data(self):
		"""Calls functions to collect original CatDV data"""
		self.ask_open_file()
		self.collect_titles()

	def add_files_to_

	def ask_save_xlsx_filename(self):
		self.xl_filename = tkFileDialog.asksaveasfilename(**self.file_options)

root = tk.Tk()
root.title('CatDV 2 CSV')
root.update()
root.minsize(root.winfo_width(), root.winfo_height())

app = CatDV2CSV(root)

root.mainloop()

