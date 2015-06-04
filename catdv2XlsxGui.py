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


class CatDV2Xlsx(tk.Frame):
	"""
	Simple program to convert a CatDV text file into an xlsx file.
	"""

	def __init__(self, parent):
		tk.Frame.__init__(self, parent)
		self.parent = parent

		self.mainf = tk.Frame(parent, bg='light yellow')
		self.open_frame = tk.Frame(self.mainf, bg='light yellow')
		self.save_frame = tk.Frame(self.mainf, bg='light yellow')

		self.mainf.grid(row=0, column=0, sticky=N+S+E+W)
		self.open_frame.grid(row=0, column=0, sticky=W+E, padx=5, pady=5) 
		self.save_frame.grid(row=1, column=0, sticky=W+E, padx=5, pady=5)	

		self.mainf.rowconfigure(0, weight=4)
		self.mainf.columnconfigure(0, weight=4)

		self.create_variables()
		self.create_menubar()
		self.create_widgets()
		self.grid_widgets()

		self.open_file_options = options = {}
		options['filetypes'] = [('text files', '.txt')]
		self.collection = []
	
	def create_variables(self):
		self.fname = tk.StringVar()
		self.xlname = tk.StringVar()

	def create_menubar(self):
		pass

	def create_widgets(self):
		self.cdv_text_btn = ttk.Button(self.open_frame, text="open CatDV text file",
			command=self.load_catdv_data, width=45)

		self.file_load = ttk.Label(self.open_frame, width=49,
			textvariable=self.fname)

		self.save_btn = ttk.Button(self.save_frame, text="save as file xlsx",
			command=self.convert_to_xlsx, width=45)

		self.save_xlsx = ttk.Label(self.save_frame, width=49,
			textvariable=self.xlname)
		self.quit_btn = ttk.Button(self.mainf, text="quit", \
			command=root.quit)

	def grid_widgets(self):
		self.cdv_text_btn.grid(row=0, column=0, columnspan=2, padx=10,
			pady=10)
		self.file_load.grid(row=1, column=0, columnspan=2, padx=10, pady=10)
		self.save_btn.grid(row=0, column=0, columnspan=2, padx=10, pady=10)
		self.save_xlsx.grid(row=1, column=0, columnspan=2, padx=10, pady=10)
		self.quit_btn.grid(row=2, column=0, sticky=W, padx=13, pady=10)

	def ask_open_file(self):
		"""Returns selected file in read mode"""
		self.cdv_file = tkFileDialog.askopenfile(
			mode='r', **self.open_file_options)
		self.fname.set("Loaded: {}".format(self.cdv_file.name))
		return self.cdv_file

	def collect_titles(self):
		"""Collects text file information into a list"""
		if self.cdv_file:
			for item in self.cdv_file:
				item.strip()
				item = item.split('\t')
				self.collection.append(item)

	def load_catdv_data(self):
		"""Calls functions to collect original CatDV data"""
		self.ask_open_file()
		self.collect_titles()

	def gen_create_xlsx(self):
		"""
		Creates an xlsx workbook for any size of text output.
		"""
		row = 0
		col = 0
		workbook = xlsxwriter.Workbook(self.xl_filename)
		self.worksheet = workbook.add_worksheet()
		for item in self.collection:
			for i in item:
				self.worksheet.write(row, col, i.rstrip())
				col += 1
				if col >= len(item):
					col = 0
					row += 1
		workbook.close()
		self.collection = []

	def ask_save_xlsx_filename(self):
		"""Ensures file is saved with a .xlsx suffix"""
		self.xl_filename = tkFileDialog.asksaveasfilename()
		if len(self.xl_filename) > 0:
			if not self.xl_filename.endswith('.xlsx'):
				self.xl_filename = self.xl_filename + '.xlsx'
		return self.xl_filename

	def convert_to_xlsx(self):

		self.ask_save_xlsx_filename()
		if len(self.xl_filename) > 0:
			self.gen_create_xlsx()
		self.xlname.set("Saved: {}".format(self.xl_filename))
		

root = tk.Tk()
root.title('CatDV 2 XLSX')
root.update()
#root.minsize(root.winfo_width(), root.winfo_height())
#root.geometry('400x200')
app = CatDV2Xlsx(root)

root.mainloop()

