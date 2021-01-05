import tkinter as tk
from tkinter.filedialog import askopenfilename


def get_file_namepath():
	root = tk.Tk()
	root.withdraw()
	filepath = askopenfilename()
	name = filepath.split('/').pop()
	path = filepath.replace(name, '')
	return name, path


def main():
	fname, fpath = get_file_namepath()	
	print('filename = {}'.format(fname))
	print('filepath = {}'.format(fpath))


if __name__ == '__main__':
	main()
