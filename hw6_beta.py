# -*- coding: utf-8 -*-
# 請參考 Moodle 平台之 "Exif 範例程式 (ReadExif3.zip)"，撰寫 GUI 的程式讀取 SEL 檔的資 料，產出 Shapefile 檔，主要工作項目為:
# (1) 從 SEL 檔 "108.sel" 擷取資料，包括:欄位 1~11。
# (2) 參考 "ShpTools.zip" 範例，產出 Shapefile 檔，及坐標定義檔 (副檔名為 PRJ)。
# (3) 程式需有圖形介面，並將程式利用 Pyinstaller 打包成為 EXE 可執行檔。
# (4) 程式執行時在視窗左上角應顯示自行設計的 LOGO 影像，可參考 "Pillow 影像處理範
# 例檔案" 將影像轉換為 ICON 影像，以節省儲存空間
import os, sys
import wx
import glob
import shapefile
# some global variables
inDir = ""
currDir = os.getcwd()
outFile = ""
data=""
dataLines=[]
#---------------------------------------------------------------------------

class MyFrame(wx.Frame):
    def __init__(self, parent, id, title):
        #create a Frame
        wx.Frame.__init__(self, parent=None, id=wx.ID_ANY, title="SEL To Shapefile Converter", size=(540,570),
                style = wx.DEFAULT_FRAME_STYLE & ~wx.MAXIMIZE_BOX ^ wx.RESIZE_BORDER)

        self.SetIcon(wx.Icon('IMG_2161.ico', wx.BITMAP_TYPE_ICO))
                
        # create a Panel
        panel = wx.Panel(self, wx.ID_ANY)
        
        wx.StaticText(parent=panel, label="輸入資料檔 ", pos=(10,10))
        self.a = wx.TextCtrl(parent=panel,pos=(90,10),size=(340,20))
        self.btn1 = wx.Button(parent=panel,label="選擇檔案",pos=(450,10),size=(80,20))
        self.Bind(wx.EVT_BUTTON, self.OnBtn1, self.btn1)

        wx.StaticText(parent=panel, label="輸出資料檔 ", pos=(10,40))
        self.b = wx.TextCtrl(parent=panel,pos=(90,40),size=(340,20))
        self.btn2 = wx.Button(parent=panel,label="選擇檔案",pos=(450,40),size=(80,20))
        self.Bind(wx.EVT_BUTTON, self.OnBtn2, self.btn2)

        self.txtCtrl = wx.TextCtrl(panel, id=wx.ID_ANY, style=wx.TE_MULTILINE, pos=(10,100), size=(520,390))

        self.btn3 = wx.Button(parent=panel,label=" Clear",pos=(10,510),size=(50,20))
        self.Bind(wx.EVT_BUTTON, self.OnBtn3, self.btn3)

        self.btn4 = wx.Button(parent=panel,label=" Convert ",pos=(465,510),size=(65,20))
        self.Bind(wx.EVT_BUTTON, self.OnBtn4, self.btn4)


    def OnBtn1(self, evt): 
        global inDir,data,dataLines
        wildcard = "*.sel" 
        dlg = wx.FileDialog(self, "Choose a file", os.getcwd(), "", wildcard, wx.FD_OPEN) 
        
        if dlg.ShowModal() == wx.ID_OK:
            try:
                with open(dlg.GetPath()) as f:
                    self.txtCtrl.WriteText("資料檔讀取成功，資料轉換中......\n")

                    path=dlg.GetPath()
                    self.a.SetValue(path) 

                    dataLines=f.readlines()

                    inDir = self.a.GetValue()
            except:
                self.txtCtrl.WriteText("輸入資料錯誤!\n")

        dlg.Destroy()

    def OnBtn2(self, evt):
        global outFile
        
        # Choose the output file. 
        dlg = wx.FileDialog(
            self, message="選擇輸出資料檔:",
            defaultDir=currDir, 
            defaultFile="",
            wildcard="*.shp",
            style= wx.FD_OVERWRITE_PROMPT | wx.FD_SAVE
            )
                     
        # If the user selects OK, then we process the dialog's data.
        # This is done by getting the path data from the dialog - BEFORE
        # we destroy it. 
        if dlg.ShowModal() == wx.ID_OK:
            # This returns a Python list of files that were selected.
            path = dlg.GetPath()
            self.b.SetValue(path)
            outFile = self.b.GetValue()      
        # Only destroy a dialog after you're done with it.
        dlg.Destroy()

    def OnBtn3(self, evt):
        
        self.txtCtrl.Clear()

    def OnBtn4(self, evt):
        try:
            result=self.shp_generate()
            if result==0:
                self.txtCtrl.WriteText('失敗!\n')
            else:
                self.txtCtrl.WriteText('成功!\n')
        except:
            self.txtCtrl.WriteText('失敗!\n')
            print('ERROR')

    def shp_generate(self):
        global data,outFile,dataLines
    
        if len(dataLines) == 0:
            return 0


        shp = shapefile.Writer(outFile[:-3], shapeType = shapefile.POINT)

        shp.field('NAME','C',10)
        shp.field('STRIP','C',10)
        shp.field('PHOTO_ID','C',40)
        shp.field('X_TWD97','N',15,3)
        shp.field('Y_TWD97','N',15,3)
        shp.field('Z_TWD97','N',15,3)
        shp.field('X_TWD67','N',15,3)
        shp.field('Y_TWD67','N',15,3)
        shp.field('Z_TWD67','N',15,3)
        shp.field('LN','C',40)
        shp.field('DATE','C',8)
        shp.field('TIME','C',20)
        shp.field('SEC','N',15,3)

        
        i=0
        for line in dataLines:   
            i+=1  
            if i <= 2:
                continue
            flight = line[:7]
            strip = line[8:10]
            photo_id = line[11:15]
            e_97 = float(line[17:27])
            n_97 = float(line[28:39])
            h_97 = float(line[40:48])
            e_67 = float(line[50:60])
            n_67 = float(line[61:72])
            h_67 = float(line[73:81])
            ln = line[82:84]
            date = line[85:91]
            time = line[94:98]
            sec = float(line[99:105])
            

            #加入點位空間資料(坐標)以及屬性資料
            
            shp.point(e_97,n_97)      #空間資料
            shp.record(flight,strip,photo_id,e_97,n_97,h_97,e_67,n_67,h_67,ln,date,time ,sec)    #屬性資料


        #存成檔名為fileOut的shapefile檔
        try:
            #關閉輸出的 shapefile 檔案才能確保資料寫入磁碟中
            shp.close()
            
            print("Shapefile圖檔產製成功!")
        except:
            print("Shapefile圖檔產製失敗!")
    


#---------------------------------------------------------------------------

# Every wxWidgets application must have a class derived from wx.App
class MyApp(wx.App):

    # wxWindows calls this method to initialize the application
    def OnInit(self):

        # Create an instance of our customized Frame class
        frame = MyFrame(None, -1, "Create Worldfile")
        frame.Show(True)

        # Tell wxWindows that this is our main window
        self.SetTopWindow(frame)

        # Return a success flag
        return True
        
#---------------------------------------------------------------------------
    
if __name__ == '__main__':
    app = MyApp(0)
    app.MainLoop()

