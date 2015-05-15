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
	Simple program to convert a CatDV text file into an xlsx file.
	"""

	def __init__(self, parent):
		tk.Frame.__init__(self, parent)
		self.parent = parent

		self.mainf = tk.Frame(parent, bg='blue')
		self.labels_frame = tk.Frame(parent, bg='gray93')
		self.btns_frame = tk.Frame(parent, bg='gray93')

		self.mainf.grid(sticky=N+S+E+W)
		self.labels_frame.grid(row=0, sticky=W+E, padx=10, pady=10)
		self.btns_frame.grid(row=0, column=1, sticky=W+E, padx=10, pady=10)	

		self.mainf.rowconfigure(0, weight=4)
		self.mainf.columnconfigure(0, weight=4)

		self.create_menubar()
		self.create_widgets()
		self.grid_widgets()

		self.open_file_options = options = {}
		options['filetypes'] = [('text files', '.txt')]
		self.collection = []


	def create_menubar(self):
		pass

	def create_widgets(self):
		self.cdv_text_label = ttk.Label(self.labels_frame, 
			text="Open CatDV text file.")
		self.cdv_text_btn = ttk.Button(self.btns_frame, text="open",
			command=self.load_catdv_data)
		self.save_label = ttk.Label(self.labels_frame,
			text="Save converted file as.")
		self.save_btn = ttk.Button(self.btns_frame, text="save and convert",
			command=self.convert_to_xlsx)
		#self.delete_text_label = ttk.Label(self.mainf, 
		#	text="Delete CatDV text file?")
		#self.delete_text_btn = ttk.Button(self.mainf, text="delete text")
		self.quit_btn = ttk.Button(self.mainf, text="quit", command=root.quit)

	def grid_widgets(self):
		self.cdv_text_label.grid(row=0, column=0)
		self.cdv_text_btn.grid(row=0, column=0)
		self.save_label.grid(row=1, column=0)
		self.save_btn.grid(row=1, column=0)
		#self.delete_text_label.grid(row=2, column=0)
		#self.delete_text_btn.grid(row=2, column=1)
		self.quit_btn.grid(row=1, column=0)

	def ask_open_file(self):
		"""Returns selected file in read mode"""
		self.cdv_file = tkFileDialog.askopenfile(
			mode='r', **self.open_file_options)
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
		print('Xlsx file created.')

	def ask_save_xlsx_filename(self):
		"""Ensures file is saved with a .xlsx suffix"""
		self.xl_filename = tkFileDialog.asksaveasfilename()
		if len(self.xl_filename) > 0:
			if not self.xl_filename.endswith('.xlsx'):
				self.xl_filename = self.xl_filename + '.xlsx'
		print self.xl_filename
		return self.xl_filename

	def convert_to_xlsx(self):

		self.ask_save_xlsx_filename()
		if len(self.xl_filename) > 0:
			self.gen_create_xlsx()
		

root = tk.Tk()
root.title('CatDV 2 XLSX')
#root.update()
#root.minsize(root.winfo_width(), root.winfo_height())
root.geometry('400x200')
app = CatDV2CSV(root)

root.mainloop()

