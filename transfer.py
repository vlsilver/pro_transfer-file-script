from PIL import Image, ImageTk
from tkinter import Tk, Label, BOTH, Text, Scrollbar,filedialog,Entry
from tkinter.ttk import Button, Style, Frame
import tkinter.messagebox as mbox
import os,sys
from os import path
import shutil
import pandas

class AppTransfer(Frame):
    #Method initial
    def __init__(self, parent):
        #Call Method khởi tạo Frame
        Frame.__init__(self,parent)
        #Dùng thuộc tính parent để lưu lại đối tượng window
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
        self.style.theme_use("default")

        global namePatchExcel
        global namePatchImage
        global patchExcel
        global patchImage
        global logText
        namePatchExcel =""
        namePatchImage =""

        frame1 = Frame(self)
        frame1.pack(side ='top',fill = BOTH)
        selectExcelButton = Button(frame1, text = "Input Excel  ",command = self.selectPatchExcel)
        selectExcelButton.pack(side = 'left')
        patchExcel = Entry(frame1,bd = 5)
        patchExcel.pack(expand = True,fill = BOTH)

        frame2 = Frame(self)
        frame2.pack(side = 'top',fill = BOTH)
        selectImageButton = Button(frame2, text = "Input Image",command = self.selectPatchImage)
        selectImageButton.pack(side = 'left')
        patchImage = Entry(frame2, bd = 5)
        patchImage.pack(expand = True,fill = BOTH)

        frame3 = Frame(self)
        frame3.pack(side = "top", pady = 10)
        doneButton = Button(frame3, text = "TRANSFER",command = self.createFolderTree)
        doneButton.pack(side = 'left', padx = 5, pady = 5)
        quitButton = Button(frame3, text = "QUIT", command=self.quit)
        quitButton.pack(side = 'right', padx = 5, pady = 5)

        logScrollbar = Scrollbar(self)
        logScrollbar.pack(side = 'right', fill = 'y')
        logText = Text(self, yscrollcommand = logScrollbar.set)
        logText.pack(side = 'bottom',fill = BOTH)
        logScrollbar.config(command = logText.yview)
    #Get Excel file patch
    def selectPatchExcel(self):
        global namePatchExcel
        namePatchExcel = filedialog.askopenfilename(title ="Select Input Excel")
        patchExcel.delete(first=0,last="end")
        patchExcel.insert('insert',namePatchExcel)
        print(namePatchExcel)
    #Get Image folder patch
    def selectPatchImage(self):
        global namePatchImage
        namePatchImage = filedialog.askdirectory(title ="Select Input Image Folder")
        patchImage.delete(first=0,last="end")
        patchImage.insert('insert',namePatchImage)
        print(namePatchImage)
    #Create Folder Tree
    def createFolderTree(self):
        logText.delete(1.0,'end')
        srcExcel = namePatchExcel
        srcImage = namePatchImage
        #Nếu file Excel tồn tại thì đọc dữ liệu
        if os.path.exists(srcExcel):
            dstImage = os.path.split(srcExcel)[0] + "\Result"
            # if os.path.exists(dstImage):
            #     shutil.rmtree(dstImage)
            # #Nếu folder file hình tồn tại thì thực hiện
            if os.path.exists(srcImage):
                #Gọi hàm tạo list Patch Image
                global listImagePatch
                global listImageName
                listImagePatch = []
                listImageName = []
                listImage = self.getPatchImage(srcImage)
                listImagePatch = listImage[0]
                listImageName = listImage[1]
                lenImageSource = len(listImageName)
                numImageTransfer = 0
                #Nếu file hình không rỗng thì
                if lenImageSource > 0:
                #Đọc dữ liệu file excel
                    excelFileImport = pandas.ExcelFile(srcExcel)
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
                            nameImage = listImageName[j]
                            #Lấy tên chương trình và mã KH trong file hình
                            if "_" in nameImage:
                                imageNameCustomerCode = nameImage.split("_")[0]
                                imageNameProgramCode = nameImage.split("_")[1]
                                #Kiểm tra xem đối tượng excel có nằm trong folder hình hay không?
                                if (checkCodeCustomer == imageNameCustomerCode) and (checkCodeProgram == imageNameProgramCode ):
                                    #Tạo đường dẫn cũ cho file hình
                                    nameImagePatch = listImagePatch[j]
                                    imageSourceOld = os.path.join(nameImagePatch,nameImage)
                                    #Tạo đường dẫn folder chứa file hình mới
                                    imageSourceNew = os.path.join(dstImage,checkProgram,checkCodeArea,checkCodeZone,checkCodeTerritory)
                                    #Đường dẫn giả định của file hình mới
                                    imageSourceNewEnd = os.path.join(imageSourceNew,nameImage)
                                    #Xóa phần từ file hình đã kiểm tra
                                    listImageName.pop(j)
                                    listImagePatch.pop(j)
                                    j -= 1
                                    lenImageSource = len(listImageName)
                                    #Nếu chưa có folder chứa hình thì tạo mới
                                    if not(os.path.exists(imageSourceNew)):
                                        #Tạo folder tree và copy
                                        os.makedirs(imageSourceNew)
                                    #Nếu file hình chưa có thì copy qua
                                    if not(os.path.exists(imageSourceNewEnd)):
                                        shutil.copy2(imageSourceOld,imageSourceNewEnd)
                                        numImageTransfer +=1
                                        logText.insert('insert','Tranfer: '+imageSourceNewEnd+"\n" )
                            j+=1
                        #Nếu đã copy hết tất cả các hình thì dừng chương trình, không cần duyệt hết file excel
                        if lenImageSource == 0:
                            break
                    _message = str("ĐÃ CHUYỂN THÀNH CÔNG: " + str(numImageTransfer))
                    mbox.showinfo("Success Transfer",_message)
                else:
                    _message = str("WARNING: ")
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
        #Duyệt qua hết tất cả các thư mục/folder
        for name in listImage:
            srcImageNew = os.path.join(_srcImage,name)
            #Kiểm tra xem nếu vẫn là thư mục thì quay lại từ đầu
            #Nếu đã là file hình thì lấy name và đường dẫn đưa vào list
            if os.path.isdir(srcImageNew) == True:
                self.getPatchImage(srcImageNew)
            else:
                listImageName.append(name)
                listImagePatch.append(_srcImage)
        return (listImagePatch,listImageName)
    

#Create a window
window = Tk()
window.geometry('600x400')
app =  AppTransfer(window)
#Display window
window.mainloop()
