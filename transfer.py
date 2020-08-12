import os
import sys
import shutil
from os import path

import pandas
import tkinter.messagebox as mbox

from PIL import Image, ImageTk
from tkinter import Tk, Label, BOTH, Text, Scrollbar,filedialog,Entry,Canvas, PhotoImage
from tkinter.ttk import Button, Style, Frame



class AppTransfer(Frame):

    #Method initial
    def __init__(self, parent):
        #Call Method khởi tạo Frame
        Frame.__init__(self,parent)
        #Dùng thuộc tính parent để lưu lại đối tượng window
        self.parent= parent

        self.namePatchExcel = ""
        self.namePatchImage = ""
        self.patchExcel = None
        self.patchImage = None
        self.logText =  None
        self.listImagePatch = []
        self.listImageName = []
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
        self.style.theme_use("default")

        # logoCanvas = Canvas(self)
        # logoCanvas.pack(side = 'top')
        # logo = ImageTk.PhotoImage(Image.open("E:\Vlprograms\Github\python-transfer_app\logo.jpg"))
        # logoCanvas.create_image(image = logo)

        frame1 = Frame(self)
        frame1.pack(side ='top',fill = BOTH)
        selectExcelButton = Button(frame1, text = "Input Excel  ",command = self.selectPatchExcel)
        selectExcelButton.pack(side = 'left')
        self.patchExcel = Entry(frame1,bd = 5)
        self.patchExcel.pack(expand = True,fill = BOTH)

        frame2 = Frame(self)
        frame2.pack(side = 'top',fill = BOTH)
        selectImageButton = Button(frame2, text = "Input Image",command = self.selectPatchImage)
        selectImageButton.pack(side = 'left')
        self.patchImage = Entry(frame2, bd = 5)
        self.patchImage.pack(expand = True,fill = BOTH)

        frame3 = Frame(self)
        frame3.pack(side = "top", pady = 10)
        doneButton = Button(frame3, text = "TRANSFER",command = self.createFolderTree)
        doneButton.pack(side = 'left', padx = 5, pady = 5)
        quitButton = Button(frame3, text = "QUIT", command=self.quit)
        quitButton.pack(side = 'right', padx = 5, pady = 5)

        logScrollbar = Scrollbar(self)
        logScrollbar.pack(side = 'right', fill = 'y')
        self.logText = Text(self, yscrollcommand = logScrollbar.set)
        self.logText.pack(side = 'bottom',fill = BOTH)
        logScrollbar.config(command = self.logText.yview)

    #Get Excel file patch
    def selectPatchExcel(self):
        #global self.namePatchExcel
        self.namePatchExcel = filedialog.askopenfilename(title ="Select Input Excel",filetypes =[("Excel file","*.xlsx"),("Excel file 97-2003","*.xls")] )
        self.patchExcel.delete(first=0,last="end")
        self.patchExcel.insert('insert',self.namePatchExcel)

    #Get Image folder patch
    def selectPatchImage(self):

        self.namePatchImage = filedialog.askdirectory(title ="Select Input Image Folder")
        self.patchImage.delete(first=0,last="end")
        self.patchImage.insert('insert',self.namePatchImage)

    #Create Folder Tree
    def createFolderTree(self):

        self.logText.delete(1.0,'end')
        #Nếu file Excel tồn tại thì đọc dữ liệu
        if os.path.exists(self.namePatchExcel):
            dstImage = os.path.split(self.namePatchExcel)[0] + "\Result"
            # if os.path.exists(dstImage):
            #     shutil.rmtree(dstImage)
            # #Nếu folder file hình tồn tại thì thực hiện
            if os.path.exists(self.namePatchImage):
                #Gọi hàm tạo list Patch Image

                self.getPatchImage(self.namePatchImage)
                print(self.listImageName)
                lenImageSource = len(self.listImageName)
                numImageTransfer = 0
                #Nếu file hình không rỗng thì
                if lenImageSource > 0:
                #Đọc dữ liệu file excel
                    excelFileImport = pandas.ExcelFile(self.namePatchExcel)
                    excelFileData = pandas.read_excel(excelFileImport,header=0,index_col=None,)
                    programExcel = list(excelFileData['Tên chương trình'].values)
                    programCodeExcel = list(excelFileData['Mã chương trình'].values)
                    areaExcel = list(excelFileData['Vùng'].values)
                    zoneExcel = list(excelFileData['Khu\nvực'].values)
                    territoryCodeExcel = list(excelFileData['Territory\nCode'].values)
                    customerCodeExcel = list(excelFileData['Mã\nKH'].values)
                    lenExcelSource = len(programCodeExcel)
                    # Kiểm tra toàn bộ các đối tượng trong file excel có trong image folder?
                    for i in range(lenExcelSource):
                        j = 0
                        checkProgram = programExcel[i]
                        checkCodeCustomer = customerCodeExcel[i]
                        checkCodeProgram = programCodeExcel[i]
                        checkCodeArea = areaExcel[i]
                        checkCodeZone = zoneExcel[i]
                        checkCodeTerritory = territoryCodeExcel[i]
                        
                        while j < lenImageSource:
                            #Tách tên file từ đường dẫn
                            nameImage = self.listImageName[j]
                            #Lấy tên chương trình và mã KH trong file hình
                            if "_" in nameImage:
                                imageNameCustomerCode = nameImage.split("_")[0]
                                imageNameProgramCode = nameImage.split("_")[1]
                                #Kiểm tra xem đối tượng excel có nằm trong folder hình hay không?
                                if (checkCodeCustomer == imageNameCustomerCode) and (checkCodeProgram == imageNameProgramCode ):
                                    #Tạo đường dẫn cũ cho file hình
                                    self.nameImagePatch = self.listImagePatch[j]
                                    imageSourceOld = os.path.join(self.nameImagePatch,nameImage)
                                    #Tạo đường dẫn folder chứa file hình mới
                                    imageSourceNew = os.path.join(dstImage,checkProgram,checkCodeArea,checkCodeZone,checkCodeTerritory)
                                    #Đường dẫn giả định của file hình mới
                                    imageSourceNewEnd = os.path.join(imageSourceNew,nameImage)
                                    #Xóa phần từ file hình đã kiểm tra
                                    self.listImageName.pop(j)
                                    self.listImagePatch.pop(j)
                                    j -= 1
                                    lenImageSource = len(self.listImageName)
                                    #Nếu chưa có folder chứa hình thì tạo mới
                                    if not(os.path.exists(imageSourceNew)):
                                        #Tạo folder tree và copy
                                        os.makedirs(imageSourceNew)
                                    #Nếu file hình chưa có thì copy qua
                                    if not(os.path.exists(imageSourceNewEnd)):
                                        shutil.copy2(imageSourceOld,imageSourceNewEnd)
                                        numImageTransfer +=1
                                        self.logText.insert('insert','Tranfer: '+imageSourceNewEnd+"\n" )
                            j+=1
                        #Nếu đã copy hết tất cả các hình thì dừng chương trình, không cần duyệt hết file excel
                        if lenImageSource == 0:
                            break
                    _message = str("ĐÃ CHUYỂN THÀNH CÔNG: " + str(numImageTransfer))
                    mbox.showinfo("Success Transfer",_message)
                else:
                    _message = str("WARNING: KHÔNG CÓ FILE HÌNH ")
                    mbox.showwarning("Warning",_message)
            else:
                _message = str("FOLDER FILE HÌNH KHÔNG TỒN TẠI")
                mbox.showerror("Error",_message)
        else:
            _message= str("EXCEL FILE NOT EXIST")
            mbox.showerror("Error",_message)

    #Get name Image and Patch of Image
    def getPatchImage(self,_srcImage):
        #Lấy list thư mục/file trong folder
        listImage = os.listdir(_srcImage)
        print(listImage)
        #Duyệt qua hết tất cả các thư mục/folder
        for name in listImage:
            print(name)
            _srcImageNew = os.path.join(_srcImage,name)
            print(_srcImage)
            #Kiểm tra xem nếu vẫn là thư mục thì quay lại từ đầu
            #Nếu đã là file hình thì lấy name và đường dẫn đưa vào list
            if os.path.isdir(_srcImageNew) == True:
                self.getPatchImage(_srcImageNew)
            else:
                self.listImageName.append(name)
                self.listImagePatch.append(_srcImage)
    

if __name__ == "__main__":
    #Create a window
    window = Tk()
    window.geometry('600x400')
    app =  AppTransfer(window)
    #Display window
    window.mainloop()
