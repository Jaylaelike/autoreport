from tkinter import *
from tkinter import filedialog
from tkinter.colorchooser import *
from tkinter.filedialog import *
root = Tk()
root.title("Auto Report Diary")
root.geometry("500x500")

def selectFile():
    # fileopen = askopenfile()
    # myLabel = Label(text=fileopen).pack()
   
    fileopen = askopenfilename()
    fileContent = open(fileopen,encoding="UTF-8")
    myLabel = Label(text=fileContent.read()).pack()


btn1 = Button(text="เลือกไฟล์",command=selectFile).pack()
root.mainloop()
