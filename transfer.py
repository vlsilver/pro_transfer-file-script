import os
import sys
import shutil
from os import path

import pandas
import tkinter.messagebox as mbox

from PIL import Image, ImageTk
from tkinter import Tk, Label, BOTH, Text, Scrollbar,filedialog,Entry,Canvas, PhotoImage, Frame, Listbox, END
from tkinter.ttk import Button, Style


class AppTransfer(Frame):

    #Method initial
    def __init__(self, parent):
        #Call Method khởi tạo Frame
        Frame.__init__(self,parent)
        #Dùng thuộc tính parent để lưu lại đối tượng window

        self.namePatchExcel =""
        self.namePatchImage =""

        self.parent= parent
        #Định nghĩa method InitUI để tạo các widget
        self.initUI()

    #Method Create UI
    def initUI(self):
        #Thay đổi tên tiêu đề cho cửa sổ
        self.parent.title("TRANFER FOLDER")
        #Định nghĩa cách sắp xếp widget lên cửa sổ
        self.pack(fill = BOTH, expand = True)
        #Quy định style cho các Widget
        self.style = Style()
        self.style.theme_use("clam")

        framePatchExcel = Frame(self)
        framePatchExcel.pack(side ='top',fill = BOTH,pady = (20,0))
        selectExcelButton = Button(framePatchExcel, text = "Input Excel     ",command = self.selectPatchExcel)
        selectExcelButton.pack(side = 'left',padx = 5)
        self.patchExcel = Entry(framePatchExcel, fg ='#00f',bd = 2)
        self.patchExcel.pack(expand = True,fill = BOTH, padx = (0,20))

        framePatchImage = Frame(self)
        framePatchImage.pack(side = 'top',fill = BOTH)
        selectImageButton = Button(framePatchImage, text = "Input Image   ",command = self.selectPatchImage)
        selectImageButton.pack(side = 'left',padx = 5)
        self.patchImage = Entry(framePatchImage, fg ='#00f', bd = 2)
        self.patchImage.pack(expand = True,fill = BOTH, padx = (0,20))

        framePatchOutput = Frame(self)
        framePatchOutput.pack(side = 'top',fill = BOTH,pady = (0,20))
        selectImageButtonOp = Button(framePatchOutput, text = "Output Image",command = self.selectPatchImageOp)
        selectImageButtonOp.pack(side = 'left',padx = 5)
        self.patchImageOp = Entry(framePatchOutput, fg ='#00f', bd = 2)
        self.patchImageOp.pack(expand = True,fill = BOTH, padx = (0,20))

        frame4 = Frame(self)
        frame4.pack(side ="top",pady = (0,20), fill = BOTH)
        checkScrollbar = Scrollbar(frame4)
        self.listBoxExcel = Listbox(frame4, fg ='#00f',height = 5, selectmode = 'multiple', yscrollcommand = checkScrollbar.set)
        self.listBoxExcel.pack(side ='left', fill = BOTH,padx = (5,0))
        checkScrollbar.pack(side = 'left', fill = 'y')
        checkScrollbar.config(command = self.listBoxExcel.yview)
        frame4_1 = Frame(frame4,height = 5)
        frame4_1.pack(side = "left")
        addPacthButton = Button(frame4_1, text = 'ADD', command = self.createPatch)
        addPacthButton.pack(side = 'top', padx = 20)
        deletePacthButton = Button(frame4_1, text = 'DELETE', command = self.deletePatch)
        deletePacthButton.pack(side = 'top', padx = 20)
        self.checkPatchText = Text(frame4, height = 5, fg ='#00f')
        self.checkPatchText.pack(side = 'left', padx = (0,20) ,pady = 0, fill = 'x', expand = True)
       
        frame3 = Frame(self)
        frame3.pack(side = "top", pady = (0,20))
        doneButton = Button(frame3, text = "TRANSFER",command = self.transfer)
        doneButton.pack(side = 'left', padx = 5, pady = 5)
        resetButton = Button(frame3, text = "RESET",command = self.reset)
        resetButton.pack(side = 'left', padx = 5, pady = 5)
        quitButton = Button(frame3, text = "QUIT", command=self.quit)
        quitButton.pack(side = 'left', padx = 5, pady = 5)


        frame5 = Frame(self)
        frame5.pack(side ='top', fill = BOTH, expand = True,pady = (0,20))
        logScrollbar = Scrollbar(frame5)
        logScrollbar.pack(side = 'right', fill = 'y')
        self.logText = Text(frame5,height =10,fg = '#000', yscrollcommand = logScrollbar.set)
        self.logText.pack(side = 'top',fill = BOTH,expand = True,padx =(5,0))
        logScrollbar.config(command = self.logText.yview)

    #STEP 1: GET ECXEL DATA
    def selectPatchExcel(self):
        #Tạo các biến
        self.excelFileData = None
        self.codeExcel = []
        self.listExcelData = []
        self.listExcelDataEnd = []
        #Xóa các dữ liệu
        self.logText.delete(1.0, END)
        self.listBoxExcel.delete('0',END)
        self.checkPatchText.delete(1.0,END)
        self.patchImageOp.delete(0,END)
        self.patchExcel.delete(0,END)
        #Chọn đường dẫn file Excel
        self.namePatchExcel = filedialog.askopenfilename(title ="Select Input Excel",filetypes =[("Excel file","*.xlsx"),("Excel file 97-2003","*.xls")] )
        #Hiện thị các đường dẫn
        self.patchExcel.insert('insert',self.namePatchExcel)
        self.namePatchImageOp = os.path.split(self.namePatchExcel)[0]
        self.patchImageOp.insert('insert', self.namePatchImageOp)
        self.checkPatchText.insert('insert', self.namePatchImageOp)
        #Đọc dữ liệu file excel
        excelFileImport = pandas.ExcelFile(self.namePatchExcel)
        self.excelFileData = pandas.read_excel(excelFileImport,header=0,index_col=None,)
        listExcelColumns = list(self.excelFileData.columns.values)
        #Bỏ các cột trống, bỏ ký tự xuống dòng
        for item in listExcelColumns:
            if not("Unname" in item):
                self.listExcelData.append(item)
                if ("\n" in item):
                    item =item.replace("\n"," ")
                self.listExcelDataEnd.append(item)
        lenExcelData = len(self.listExcelData)
        #Hiện thị các cột excel để add đường dẫn
        for i in range(lenExcelData):
            self.listBoxExcel.insert(i,self.listExcelDataEnd[i])
        self.logText.insert('insert',"STEP 1: GET EXCEL DATA\nNote that: Output Image patch ís in same folder of Excel File (Click 'Output Image' button to Change!)")

    #STEP 2: GET IMAGE DATA
    def selectPatchImage(self):
        #Yêu cầu thực hiện bước 1
        if len(self.namePatchExcel) > 0:
            #Đặt các biến
            self.listImageName = []
            self.listImagePatch = []
            #Chọn đường dẫn folder hình
            self.namePatchImage = filedialog.askdirectory(title ="Select Input Image Folder")
            self.patchImage.insert('insert',self.namePatchImage)
            #Gọi hàm để lấy dữ liệu hình
            self.getPatchImage(self.namePatchImage)
            if len(self.listImageName) > 0:
                self.logText.insert('insert',"\n\nSTEP 2: CREATE IMAGE LIST \n"+str(len(self.listImageName))+" file are founded")
            else:
                mbox.showinfo("Warning","Folder is nothing!")
        else:
            mbox.showinfo("Warning","Please fill Input Excel File patch \nClick 'Input Excel' Button")

    #STEP 2: GET IMAGE DATA, CALLED BY selectPatchImage()
    def getPatchImage(self,_srcImage):
        listImage = os.listdir(_srcImage)
        for name in listImage:
            _srcImageNew = os.path.join(_srcImage,name)
            if os.path.isdir(_srcImageNew) == True:
                self.getPatchImage(_srcImageNew)
            else:
                self.listImageName.append(name)
                self.listImagePatch.append(_srcImage)
   
    #STEP 3: CREATE OUTPUT IMAGE FOLDER
    def createPatch(self):
        #Kiểm tra xem STEP 1 vs 2 đã thực hiện chưa
        if len((self.namePatchExcel)) > 0 and len(self.namePatchImage) > 0:
            #Lấy list index item được chọn
            index = self.listBoxExcel.curselection()
            #Tạo đường dẫn cho người dùng kiểm tra
            for i in index:
                self.codeExcel.append(self.listExcelData[i])
                textPatch = os.path.join("/",self.listBoxExcel.get(i))
                self.checkPatchText.insert('insert',textPatch)
                self.patchImageOp.insert('insert',textPatch)
            #Reset lại item đã chọn
            self.listBoxExcel.delete('0',END)
            lenExcelData = len(self.listExcelData)
            for i in range(lenExcelData):
                self.listBoxExcel.insert(i,self.listExcelDataEnd[i])
            self.logText.insert("insert","\n\nSTEP 3: CREATE OUPUT IMAGE PATCH\nOutput Image Patch Current: "+self.checkPatchText.get(1.0,END))
        else:
            mbox.showinfo("Warning","Please fill Input Excel, Input Image  File patch \nClick 'Input Excel', 'Input Image' Button")

    #STEP 3: CREATE OUTPUT IMAGE FOLDER
    def deletePatch(self):
        if len(self.namePatchExcel) > 0:
            self.codeExcel = []
            self.checkPatchText.delete(1.0,END)
            self.patchImageOp.delete(first=0,last="end")
            self.patchImageOp.insert('insert',os.path.split(self.namePatchExcel)[0])
            self.checkPatchText.insert('insert',os.path.split(self.namePatchExcel)[0])

    #STEP 3: CREATE OUTPUT IMAGE FOLDER
    def selectPatchImageOp(self):
        if len(self.namePatchExcel) > 0:
            self.codeExcel= []
            self.checkPatchText.delete(1.0,END)
            self.namePatchImageOp = filedialog.askdirectory(title ="Select Input Image Folder")
            self.patchImageOp.delete(0,END)
            self.patchImageOp.insert('insert',self.namePatchImageOp)
            self.checkPatchText.insert('insert',self.namePatchImageOp)
            lenExcelData = len(self.listExcelData)
            for i in range(lenExcelData):
                self.listBoxExcel.insert(i,self.listExcelDataEnd[i])
        else:
            mbox.showinfo("Warning","Please fill Input Excel File patch \nClick 'Input Excel' Button")

    #STEP 4: TRANSFER IMAGE FILE
    def transfer(self):
        key_1 = "Mã chương trình"
        key_2 = "Mã\nKH"
        key_split = "_"
        if len(self.namePatchExcel) > 0 and len(self.namePatchImage) > 0:
            codeExcelData = list(self.codeExcel)
            listImageNameData = list(self.listImageName)
            listImagePatchData = list(self.listImagePatch)
            self.logText.insert('insert',"\nSTEP 4: TRANSFER FILE WITH KEY")
            lenImageSource = len(listImageNameData)
            lenCode = len(self.codeExcel)
            numImageTransfer = 0
            if lenImageSource > 0:
                if lenCode > 0:
                    for i in range(lenCode):
                        codeExcelData[i] = list(self.excelFileData[self.codeExcel[i]].values)
                    key_1_list = list(self.excelFileData[key_1].values)
                    key_2_list = list(self.excelFileData[key_2].values)
                    lenExcelSorce = len(key_2_list)
                    for i in range(lenExcelSorce):
                        key_1_excel = key_1_list[i]
                        key_2_excel = key_2_list[i]
                        j = 0
                        while j < lenImageSource:
                            nameImage = listImageNameData[j]
                            if key_split in nameImage:
                                key_2_image = nameImage.split("_")[0]
                                key_1_image = nameImage.split("_")[1]
                                if (key_1_excel == key_1_image) and (key_2_excel == key_2_image):
                                    nameImagePatch = listImagePatchData[j]
                                    imageSourceOld = os.path.join(nameImagePatch,nameImage)
                                    imageSourceNew = self.namePatchImageOp
                                    for k in range(lenCode):
                                        imageSourceNew = os.path.join(imageSourceNew,codeExcelData[k][i])
                                    imageSourceNewEnd = os.path.join(imageSourceNew,nameImage)
                                    listImageNameData.pop(j)
                                    listImagePatchData.pop(j)
                                    j -= 1
                                    lenImageSource = len(listImageNameData)
                                    if not(os.path.exists(imageSourceNew)):
                                        os.makedirs(imageSourceNew)
                                    if not(os.path.exists(imageSourceNewEnd)):
                                        shutil.copy2(imageSourceOld,imageSourceNewEnd)
                                        numImageTransfer +=1
                                        self.logText.insert('insert','    Tranfer: '+imageSourceNewEnd+"\n" )
                            j+=1
                        if lenImageSource == 0:
                            break
                    _message = str("ĐÃ CHUYỂN THÀNH CÔNG: " + str(numImageTransfer))
                    mbox.showinfo("Success Transfer",_message)
                else:
                   mbox.showinfo("Warning","Please check your Output Image Patch") 
            else:
                mbox.showinfo("Warning","Image Folder is nothing")
        else:
            mbox.showinfo("Warning","Please fill Input Excel File, Input Image File patch \nClick 'Input Excel', 'Input Image' Button")

    #Reset
    def reset(self):
        self.namePatchExcel =[]
        self.patchExcel.delete(0,END)
        self.namePatchImage = []
        self.patchImage.delete(0,END)
        self.namepatchImageOp = []
        self.patchImageOp.delete(0,END)
        self.listBoxExcel.delete(0,END)
        self.checkPatchText.delete(1.0,END)
        self.logText.delete(1.0,END)
        mbox.showwarning("Warning","You just RESET!")
if __name__ == "__main__":
    #Create a window
    window = Tk()
    window.geometry('700x500')
    app =  AppTransfer(window)
    #Display window
    window.mainloop()
