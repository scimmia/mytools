from tkinter.filedialog import askopenfilename

def selectFile(pathFile,filetypes=[('XLS', '*.xls;*.xlsx'), ('All Files', '*')]):
    file_path = askopenfilename(filetypes=filetypes)
    pathFile.set(file_path)