import tkinter as tk
import tkinter.font as tkFont
from tkinter import ttk
from tkinter import filedialog
import os
from pdfrw import PdfReader, PdfWriter
import pyautogui
import datetime

version = "_ver1.0"

class Application():
    def __init__(self, master):
        self.master = master
        self.foldername = None
        self.pdf_merge_output = None
        '''
        window
        '''
        window.geometry('330x330+1200+600')  # set new geometry , window appear at (1200, 600)
        window.title(("PDF文書工具" + version))
        window.resizable(width=False, height=False)
        #TK font
        self.font1 = tkFont.Font(family='Times', size=16, weight='bold', underline=0)
        self.font2 = tkFont.Font(family='Times', size=14, weight='bold', underline=0)
        self.font3 = tkFont.Font(family='Times', size=12)
        self.font4 = tkFont.Font(family='Times', size=10)
        #TK frame
        self.frame1 = tk.Frame(master=window, width=330, height=110)
        self.frame2 = tk.Frame(master=window, width=330, height=110)
        self.frame3 = tk.Frame(master=window, width=330, height=110)
        self.frame1.pack()
        self.frame2.pack()
        self.frame3.pack()
        '''
        label & message
        '''
        self.lb_pdf_merge = tk.Label(master=self.frame1,text="PDF合併",font=self.font1).place(x=10, y=0)
        self.mg_pdf_merge_path = tk.Message(master=self.frame1,text="瀏覽PDF合併資料夾，將依檔名順序合併",
                            font=self.font4, width= 300, bd=2, fg='Azure4',padx=1,anchor= "nw")
        self.mg_pdf_merge_path.place(x=10, y=25)
        self.lb_pdf_merge_count = tk.Label(master=self.frame1, text ="共0個PDF檔",font=self.font4)
        self.lb_pdf_merge_count.place(x=120, y=85)
        self.lb_pdf_rotate = tk.Label(master=self.frame2,text="PDF旋轉",font=self.font1).place(x=10, y=0)
        self.mg_pdf_rotate_path = tk.Message(master=self.frame2,text="選擇PDF旋轉檔案",
                            font=self.font4, width= 300, bd=2, fg='Azure4',padx=1,anchor= "nw")
        self.mg_pdf_rotate_path.place(x=10, y=25)
        self.lb_pdf_rotate_cloclwise = tk.Label(master=self.frame2,text="順時針轉",font=self.font4).place(x=100, y=85)
        
        self.lb_pdf_remove = tk.Label(master=self.frame3,text="PDF移除指定頁",font=self.font1).place(x=10, y=0)
        self.mg_pdf_remove = tk.Message(master=self.frame3,text="選擇檔案移除PDF指定頁",
                                    font=self.font4, width= 300, bd=2, fg='Azure4',padx=1,anchor= "nw")
        self.mg_pdf_remove.place(x=10, y=25)
        self.lb_pdf_remove_count = tk.Label(master=self.frame3, text ="共0頁 輸入刪除頁:",font=self.font4)
        self.lb_pdf_remove_count.place(x=70, y=80)
        self.entry_pdf_remove_page = tk.Entry(master=self.frame3, text ="",font=self.font4, width = 8)
        self.entry_pdf_remove_page.place(x=200, y=80)
        '''
        button
        '''
        # #button browse pdf merge folder path
        self.btn_pdf_merge_path = tk.Button(master = self.frame1, text="瀏覽..", font=self.font4, 
                                    command=lambda:[self.browse_pdf_merge_folder(),self.update_widget1()])
        self.btn_pdf_merge_path.place(x=20, y=80)
        #button excute pdf merge
        self.btn_pdf_merge = tk.Button(master = self.frame1, text=" 執行 ", font=self.font4,
                                    command = self.pdf_merge, state = "disabled")
        self.btn_pdf_merge.place(x=270, y=80)

        #button browse pdf rotate path
        self.btn_pdf_rotate_path = tk.Button(master = self.frame2, text="瀏覽..", font=self.font4,
                                    command=self.browse_pdf_file)
        self.btn_pdf_rotate_path.place(x=20, y=80)
        #selection menu for rotation angle
        self.angle_menu = ttk.Combobox(master = self.frame2, state="readonly", width=6,
                            values=["90度", 
                                    "180度",
                                    "270度",])
        self.angle_menu.current(0)#default select 90 degree
        self.angle_menu.place(x=175, y=85)                                                
        #button excute pdf rotate
        self.btn_pdf_rotate = tk.Button(master = self.frame2, text=" 執行 ", font=self.font4,
                                command= self.pdf_rotate, state = "disabled")
        self.btn_pdf_rotate.place(x=270, y=80)
        # #button browse pdf merge folder path

        self.btn_pdf_remove_path = tk.Button(master = self.frame3, text="瀏覽..", font=self.font4,
                                    command= lambda:[self.browse_pdf_remove_file(), self.update_widget2()])
        self.btn_pdf_remove_path.place(x=20, y=80)
        #button excute pdf merge
        self.btn_pdf_remove = tk.Button(master = self.frame3, text=" 執行 ", font=self.font4,
                                    command = self.pdf_remove, state = "disabled")
        self.btn_pdf_remove.place(x=270, y=80)

        '''
        method
        '''
    def browse_pdf_merge_folder(self):
        self.foldername = filedialog.askdirectory(title='瀏覽PDF合併資料夾')
        if self.foldername:
            self.btn_pdf_merge['state'] = 'normal'
            try:
                self.mg_pdf_merge_path.config(fg='black', text = self.foldername)    
                #break
            except FileExistsError:
                self.mg_pdf_merge_path.config(fg='black', text = self.foldername)

    def browse_pdf_file(self):
        self.filename = filedialog.askopenfilename(filetypes = [("PDF files", "*.pdf")],
                                        title='選擇PDF旋轉檔案')
        if self.filename:
            self.btn_pdf_rotate['state'] = 'normal'
            self.excel_file = self.filename
            self.mg_pdf_rotate_path.config(fg='black', text = self.filename)

    def browse_pdf_remove_file(self):
        self.filename = filedialog.askopenfilename(filetypes = [("PDF files", "*.pdf")],
                                        title='選擇移除PDF空白頁檔案')
        if self.filename:
            self.btn_pdf_remove['state'] = 'normal'
            self.excel_file = self.filename
            self.mg_pdf_remove.config(fg='black', text = self.filename)


    def files_list(self, folder_name, file_type):
        self.files_name=[]
        for self.filename in os.listdir(folder_name):
            if self.filename.endswith(str(file_type)): 
                self.files_name.append(self.filename)
        return self.files_name

    def update_widget1(self):
        #print("text update")
        if self.foldername ==None:
            print("Folder havn't set")
            pass
        else:
            self.lb_pdf_merge_count["text"] = "共 "+str(len(self.files_list(self.foldername, ".pdf")))+" 個PDF檔"
            window.update_idletasks()

    def update_widget2(self):
        #print("text update")
            self.pdf_being_remove = PdfReader(self.filename)
            self.lb_pdf_remove_count["text"] = "共"+str(len(self.pdf_being_remove.pages))+"頁 輸入刪除頁:"
            window.update_idletasks()


    def pdf_merge(self):
        self.pdf_files = self.files_list(self.foldername, ".pdf")
        self.timestamp = datetime.datetime.today().strftime("%Y%m%d%H%M")
        self.pdf_merge_output = self.foldername + "/merge"
        if len(self.files_list(self.foldername, ".pdf"))<=1:
            pyautogui.alert("資料夾內無PDF檔案或僅一個PDF檔")
        else:
            try:
                os.makedirs(self.pdf_merge_output)
            except:
                pass
            self.output_pdf = self.pdf_merge_output + "\\"  + self.pdf_files[0].replace(".pdf","") + "_merge_"+ self.timestamp + ".pdf"
            self.writer = PdfWriter()
            for in_pdf in self.pdf_files:
                self.writer.addpages(PdfReader(self.foldername + "\\" + in_pdf).pages)

            self.writer.write(self.output_pdf)
            pyautogui.alert("PDF合併完成，請確認。")


    def pdf_rotate(self):
        #get combox value for rotation angle 
        if self.angle_menu.get() == '90度':
            self.angle = 90
        elif self.angle_menu.get() == '180度':
            self.angle = 180
        elif self.angle_menu.get() == '270度':
            self.angle = 270
        try:
            self.trailer = PdfReader(self.filename)
            self.pages = self.trailer.pages
            #rotate all page in pdf file
            self.ranges = [[1, len(self.pages)]]
            for onerange in self.ranges:
                onerange = (onerange + onerange[-1:])[:2]
                for pagenum in range(onerange[0]-1, onerange[1]):
                    self.pages[pagenum].Rotate = (int(self.pages[pagenum].inheritable.Rotate or
                                                0) + self.angle) % 360
            #get pdf file folder path = filepath - filename
            self.output_pdf = self.filename.replace(os.path.basename(self.filename),"")
            #extract the path without file extention
            self.temp_filename = os.path.splitext(self.filename)[0]
            #output file name = filename_merge.pdf
            self.pdf_rotate_filename = os.path.basename(self.temp_filename) + '_' +str(self.angle) + ".pdf"
            #output file = folder path + filename_merge.pdf
            self.outdata = PdfWriter(self.output_pdf + self.pdf_rotate_filename)
            self.outdata.trailer = self.trailer
            self.outdata.write()
            pyautogui.alert("PDF旋轉完成，請確認。")
        except:
            pyautogui.alert("檔案開啟錯誤")

    def pdf_remove(self):
        #get folder path = full file path - file name
        self.output_rm_path = self.filename.replace(os.path.basename(self.filename),"") + "/remove"
        if len(self.pdf_being_remove.pages) <= 1:
            pyautogui.alert("PDF檔案僅有一頁")
        else:
            self.output_rm_pdf = self.output_rm_path + "\\" + os.path.basename(self.filename).replace(".pdf","")\
            + "_removed_" + ".pdf"
            self.rm_page_raw = str(self.entry_pdf_remove_page.get()).split(",")
            try:
                os.makedirs(self.output_rm_path)
            except:
                pass
            #processing: check if user input valid 
            self.rm_page = []
            self.error_flag = 0
            for i in self.rm_page_raw:
                try:
                    self.rm_page.append(int(i)-1)
                    self.rm_page.sort()
                    if int(i) > len(self.pdf_being_remove.pages) or int(i)<1:
                        pyautogui.alert("刪除頁輸入錯誤:超過檔案頁數")
                        self.error_flag = 1
                        break
                except ValueError:
                    pyautogui.alert("刪除頁輸入錯誤:非數字或特殊符號，多頁以"",""分隔")
                    self.error_flag = 1
                    break
            self.pdf_rm_output = PdfWriter()
            if self.error_flag == 0:
                for current_page in range(len(self.pdf_being_remove.pages)):
                    if current_page in self.rm_page:
                        print("Page being removed:", current_page+1)
                        pass
                    else:
                        self.pdf_rm_output.addpage(self.pdf_being_remove.pages[current_page])
                        print("adding page %i" % (current_page+1))
                self.pdf_rm_output.write(self.output_rm_pdf)
                pyautogui.alert("指定頁面移除完成，請確認。")

if __name__ == '__main__':
    window = tk.Tk()
    app = Application(window)
    window.mainloop()
