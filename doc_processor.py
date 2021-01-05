import tkinter as tk
from tkinter.filedialog import askopenfilename
from docx import Document


def get_file_namepath():
	""" Open modal dialog to get file location from user
	and return file name and file path"""
	root = tk.Tk()
	root.withdraw()
	filepath = askopenfilename()
	name = filepath.split('/').pop()
	path = filepath.replace(name, '')
	return name, path, filepath


def process_doc(filename):
	""" Crawl through the file and add html tags """
	document = Document(filename)
	return document


def main():
	fname, fpath, fp = get_file_namepath()
	doc = process_doc(fp)

	# save updated document to same directory as original file
	doc.save(fpath + 'PROCESSED_' + fname)


if __name__ == '__main__':
	main()
