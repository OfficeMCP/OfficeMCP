import winreg
import asyncio
import pywintypes
import os,sys
import win32com.client
from pathlib import Path
from fastmcp.resources import FileResource, TextResource, DirectoryResource
class TheOfficer:
    def __init__(self):
        self._excel = None
        self._word = None
        self._outlook = None
        self._visio =None
        self._access = None
        self._project = None
        self._publisher = None
        self._onenote = None
        self._powerpoint = None
        self._ket = None
        self._kwpp = None
        self._kwps = None

        self._default_folder ="D:\\@OfficeMCP"

        self.Version = "1.0.4"

        self.Speaker = None

        self.ComObjects={}

        self._printable = False
        self.Data = OfficerData()

        self.MicrosoftApplications=[
        'Word', 'Excel', 'PowerPoint',
        'Visio', 'Access', 'MSProject',
        'Outlook', 'Publisher',"OneNote","Kwps","Ket","Kwpp"
        ]
    
    @property
    def RootFolder(self)->str:
        '''created if not exists'''
        if not os.path.exists(self._default_folder):
            os.makedirs(self._default_folder)
        return self._default_folder
    
    def GetComObject(self,com_name:str,active:bool=True)->bool:
        """Speak the text. volume range is 0-100, rate range is -10 to 10."""
        comobject = self.ComObjects.get(com_name)
        if comobject is None:
            try:
                comobject=win32com.client.GetActiveObject(com_name)
            except Exception as e:
                comobject=win32com.client.Dispatch(com_name)
            self.ComObjects[com_name]=comobject
        return comobject

    def Speak(self,text: str="Hello!, I'm Office MCP.", volume: int = 80, rate: int = 2)->bool:
        """Speak the text. volume range is 0-100, rate range is -10 to 10."""
        try:
            if self.Speaker is None:
                self.Speaker = self.GetComObject("SAPI.SPVOICE")
            self.Speaker.Volume = volume
            self.Speaker.Rate = rate
            self.Speaker.Speak(text)
            return True
        except Exception as e:
            print(e)
            return False
    
    def Beep(self,frequency:int=500,duration:int=500):
        import winsound
        winsound.Beep(frequency, duration) 

    def DownloadImage(self, url: str, save_path: str = None) -> str:
        """Download an image from the given URL and save it to the specified path under the OfficeMCP root folder."""
        import datetime
        if save_path is None:
            now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            save_path = f"downloaded_{now}.png"
        return self.DownloadURL(url, save_path)
        
    def DownloadURL(self, url: str, save_path: str = None) -> str:
        """Download a URL content from the given URL and save it to the specified path."""
        import urllib.request
        import datetime
        if save_path is None:
            now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            save_path = f"downloaded_{now}.png"
        file_path = self.FilePath(save_path)
        try:
            with urllib.request.urlopen(url) as response, open(file_path, 'wb') as out_file:
                out_file.write(response.read())
            return file_path
        except Exception as e:
            print(f'Download failed: {e}')
            return ""
        
    def Print(self, obj):
        if self._printable:
            try:              
                print(obj)
            except Exception as e:
                print(e)

    def Visible(self, app_name: str, visible = None) -> bool:
        """Check if the specified application is visible."""
        try:
            app = self.Application(app_name)
            if app is None:
                return False
            else:
                if not visible is None:
                    app.Visible = visible
                return app.Visible
        except Exception as e:
            print(e)
            return False

    def Quit(self,app_name: str, force: bool = False)->bool:
        """Quit the microsoft excel application."""
        app_name_attr="_"+app_name.lower()
        if hasattr(self, app_name_attr):
            app = getattr(self, app_name_attr)
            try:
                return self.QuitApplication(app, force)
            except Exception as e:
                print(e)
        return False

    def QuitApplication(self,app,force: bool = False)->bool:
        if not force:
            try:
                app.Quit()
                return True
            except Exception as e:
                print(e)
                return False
        import win32process
        import win32api
        import win32con
        # Get the window's process id's
        hwnd = app.Hwnd
        t, p = win32process.GetWindowThreadProcessId(hwnd)
        # Ask window nicely to close  
        try:
            handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, 0, p)
            if handle:
                win32api.TerminateProcess(handle, 0)
                win32api.CloseHandle(handle)
                return True
        except:
            pass
        return False

    def ComApplication(self, com_diy_name: str, com_full_name,asNewInstance: bool = False) -> object:  # noqa: E741
        """Get the specified Microsoft Office application object."""
        app_name_attr="_"+com_diy_name.lower()
        if hasattr(self, app_name_attr):
            app = getattr(self, app_name_attr)
            try:
                return app
            except Exception as e:
                print(e)
        if asNewInstance:
            app = win32com.client.Dispatch(com_full_name)
            self.__dict__[app_name_attr] = app
            return app
        else:
            try:
                app = win32com.client.GetActiveObject(com_full_name)
            except pywintypes.com_error:
                app = win32com.client.Dispatch(com_full_name)
            self.__dict__[app_name_attr] = app
            return app

    def Application(self, app_name: str, asNewInstance: bool = False) -> object:  # noqa: E741
        """Get the specified Microsoft Office application object."""
        app_name_attr="_"+app_name.lower()
        if hasattr(self, app_name_attr):
            app = getattr(self, app_name_attr)
            try:
                name = app.Name
                return app
            except Exception as e:                
                #print(e)
                pass
        if not app_name in self.MicrosoftApplications:
            return None
        if not self.IsAppAvailable(app_name):
            return None
        app_full_name = app_name + ".Application"#
        if asNewInstance:
            app = win32com.client.Dispatch(app_full_name)
            self.__dict__[app_name_attr] = app
            return app
        else:
            try:
                app = win32com.client.GetActiveObject(app_full_name)
            except pywintypes.com_error:
                app = win32com.client.Dispatch(app_full_name)
            self.__dict__[app_name_attr] = app
            return app

    def RunningApps(self) -> list:
        apps = []
        for prog_id in self.MicrosoftApplications:
            try:
                app = win32com.client.GetActiveObject(prog_id + ".Application")
                if app is not None:
                    apps.append(prog_id)
            except (FileNotFoundError, pywintypes.com_error):
                continue
            except Exception as e:
                print(f"[DEBUG] Error verifying {prog_id}: {str(e)}")
        return apps

    def AvailableApps(self) -> list:        
        apps = []
        for prog_id in self.MicrosoftApplications:
            try:
                # Registry verification
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, prog_id+".Application"):
                    pass
                apps.append(prog_id)            
            except (FileNotFoundError, pywintypes.com_error):
                continue
            except Exception as e:
                print(f"[DEBUG] Error verifying {prog_id}: {str(e)}")            
        return apps
    def IsAppAvailable(self,app_name: str) -> bool:
        """Check if the specified application is installed."""
        try:
            if not app_name.endswith(".Application"):
                app_name=app_name+".Application"
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, app_name):
                pass
            return True
        except Exception as e:
            print("e")
            return False

    def Demonstrate(self)->str:
        if not self.DemonstrateExcel():
            return  "There's Excel installed or something wrong"
        if not self.DemonstratePowerPoint():
            return "There's no PowerPoint installed or something wrong"
        return "Demo succesed"
    def DemonstrateExcel(self)->bool:
        excel=self.Excel
        if excel is None:
            return  False
        book = excel.Workbooks.Add()
        sheet = excel.ActiveSheet
        excel.Visible = True
        sheet.Cells(1, 1).Value = "Hello, World From OfficeMCP Server!"
        sheet.Cells(1, 1).Font.Size = 20
        sheet.Cells(1, 1).Font.Bold = True
        sheet.Cells(2, 1).Value = "This is a demonstration of the OfficeMCP server."
        sheet.Cells(3, 1).Value = "You can use this server's tool to control all Microsoft Office Applications."
        sheet.Cells(4, 1).Value = "RunPython tool to run Python codes to edit Office Application's documents."
        sheet.Cells(4, 1).Font.Color = 0xFF0000  # Red color
        sheet.Cells(5, 1).Value = "AvailableApplications tool to check the Microsoft Office applications installed."
        sheet.Cells(6, 1).Value = "IsAppAvailable tool to check if an specific Microsoft Office application is installed."
        sheet.Cells(7, 1).Value = "Launch tool to Launch an specific Microsoft Office application."
        sheet.Cells(8, 1).Value = "Visible tool to check if an specific Microsoft Office application is visible."
        sheet.Cells(9, 1).Value = "Quit tool to quit an specific Microsoft Office application."  
        return True      
    def DemonstratePowerPoint(self)->bool:
        ppt=self.Application("PowerPoint")
        if ppt is None:
            return False
        ppt.Visible = True
        presentation = ppt.Presentations.Add()
        slide = presentation.Slides.Add(1, 12)
        top = 20
        shape = slide.Shapes.AddTextbox(1, 100, top, 800, 100)
        shape.TextFrame.TextRange.Text = "Hello, World From OfficeMCP Server!"
        #make shape bold
        shape.TextFrame.TextRange.Font.Bold = True
        top += 40
        shape = slide.Shapes.AddTextbox(1, 100, top, 800, 100)
        shape.TextFrame.TextRange.Text = "This is a demonstration of the OfficeMCP server."
        top += 40
        shape = slide.Shapes.AddTextbox(1, 100, top, 800, 100)
        shape.TextFrame.TextRange.Text = "You can use this server's tools to control all Microsoft Office Applications."
        top += 40
        shape = slide.Shapes.AddTextbox(1, 100, top, 800, 100)
        shape.TextFrame.TextRange.Text = "RunPython tool to run Python codes to edit Office Application's documents."
        #make shape red
        shape.TextFrame.TextRange.Font.Color = 0xFF0000
        top += 40
        shape = slide.Shapes.AddTextbox(1, 100, top, 800, 100)
        shape.TextFrame.TextRange.Text = "AvailableApplications tool to check the Microsoft Office applications installed."
        top += 40
        shape = slide.Shapes.AddTextbox(1, 100, top, 800, 100)
        shape.TextFrame.TextRange.Text = "IsAppAvailable tool to check if an specific Microsoft Office application is installed."
        top += 40
        shape = slide.Shapes.AddTextbox(1, 100, top, 800, 100)
        shape.TextFrame.TextRange.Text = "Launch tool to Launch an specific Microsoft Office application."
        top += 40
        shape = slide.Shapes.AddTextbox(1, 100, top, 800, 100)
        shape.TextFrame.TextRange.Text = "Visible tool to check if an specific Microsoft Office application is visible."    
        top += 40
        shape = slide.Shapes.AddTextbox(1, 100, top, 800, 100)
        shape.TextFrame.TextRange.Text = "Quit tool to quit an specific Microsoft Office application."
        top += 40
        shape = slide.Shapes.AddTextbox(1, 100, top, 800, 100)
        shape.TextFrame.TextRange.Text = "Demonstrate tool to see this demonstration."
        top += 40
        shape = slide.Shapes.AddTextbox(1, 100, top, 800, 100)
        shape.TextFrame.TextRange.Text = "Instructions tool to see the instructions of this server."
        return True

    def FilePath(self, file_name: str=None, subfolder:str = None) -> str:
        if subfolder is None:
            file_path =os.path.join(self.RootFolder,file_name)
        else:
            file_path = os.path.join(self.RootFolder, subfolder,file_name)
        return file_path

    def IsFileExits(self, file_name: str) -> bool:
        """Check if the specified file exists."""
        file_path = self.FilePath(file_name)
        return os.path.exists(file_path)
    
    def ScreenShot(self, save_path: str = None) -> str:
        """Capture a screenshot of the entire screen and save to the specified path."""
        import win32gui
        import win32ui
        import win32con
        import win32api
        from PIL import Image
        import datetime
        # 获取桌面窗口句柄
        hdesktop = win32gui.GetDesktopWindow()
        width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
        height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
        left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
        top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
        desktop_dc = win32gui.GetWindowDC(hdesktop)
        img_dc = win32ui.CreateDCFromHandle(desktop_dc)
        mem_dc = img_dc.CreateCompatibleDC()
        screenshot = win32ui.CreateBitmap()
        screenshot.CreateCompatibleBitmap(img_dc, width, height)
        mem_dc.SelectObject(screenshot)
        mem_dc.BitBlt((0, 0), (width, height), img_dc, (left, top), win32con.SRCCOPY)
        bmpinfo = screenshot.GetInfo()
        bmpstr = screenshot.GetBitmapBits(True)
        img = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)
        if save_path is None:
            now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            save_path =self.FilePath(f"screenshot_{now}.png")
        else:
            save_path =self.FilePath(save_path)
        img.save(save_path)
        # 释放资源
        mem_dc.DeleteDC()
        win32gui.DeleteObject(screenshot.GetHandle())
        img_dc.DeleteDC()
        win32gui.ReleaseDC(hdesktop, desktop_dc)
        return save_path

 #region Properties

    @property
    def Kwps(self) -> object:
        """Get the WPS word application object."""
        try:
            name = self._kwps.Name
            return self._kwps
        except Exception as e:
            self._kwps = self.Application("Kwps",False)
        return self._kwps
    
    @property
    def Ket(self) -> object:
        """Get the WPS Excel application object."""
        try:
            name = self._ket.Name
            return self._ket
        except Exception as e:
            self._ket = self.Application("Ket",False)
        return self._ket
    
    @property
    def Kwpp(self) -> object:
        """Get the WPS powerpoint application object."""
        try:
            name = self._kwpp.Name
            return self._kwpp
        except Exception as e:
            self._kwpp = self.Application("Kwpp",False)
        return self._kwpp

    @property
    def Excel(self) -> object:
        """Get the Excel application object."""
        try:
            name = self._excel.Name
            return self._excel
        except Exception as e:
            self._excel = self.Application("Excel",False)
        return self._excel

    @property
    def Word(self) -> object:
        """Get the Word application object."""
        try:
            name = self._word.Name
            return self._word
        except Exception as e:
            self._word = self.Application("Word",False)
        return self._word
    
    @property
    def PowerPoint(self) -> object:
        """Get the PowerPoint application object."""
        try:
            name = self._powerpoint.Name
            return self._powerpoint
        except Exception as e:
            self._powerpoint = self.Application("PowerPoint", False)
        return self._powerpoint

    @property
    def Visio(self) -> object:
        """Get the Visio application object."""
        try:
            name = self._visio.Name
            return self._visio
        except Exception as e:
            self._visio = self.Application("Visio", False)
        return self._visio

    @property
    def Access(self) -> object:
        """Get the Access application object."""
        try:
            name = self._access.Name
            return self._access
        except Exception as e:
            self._access = self.Application("Access", False)
        return self._access

    @property
    def Project(self) -> object:
        """Get the Project application object."""
        try:
            name = self._project.Name
            return self._project
        except Exception as e:
            self._project = self.Application("MSProject", False)
        return self._project

    @property
    def Outlook(self) -> object:
        """Get the Outlook application object."""
        try:
            name = self._outlook.Name
            return self._outlook
        except Exception as e:
            self._outlook = self.Application("Outlook", False)
        return self._outlook

    @property
    def Publisher(self) -> object:
        """Get the Publisher application object."""
        try:
            name = self._publisher.Name
            return self._publisher
        except Exception as e:
            self._publisher = self.Application("Publisher", False)
        return self._publisher

    @property
    def OneNote(self) -> object:
        """Get the OneNote application object."""
        try:
            name = self._onenote.Name
            return self._onenote
        except Exception as e:
            self._onenote = self.Application("OneNote", False)
        return self._onenote

 #endregion   

class OfficerData:
    def init(self):
        # Get current attributes
        current_attrs = set(vars(self).keys())       
        # Delete attributes not in the initial set
        for attr in current_attrs - self._initial_attrs:
            delattr(self, attr)    
        self.data ={}
        self.output=""
        self.helloworld = "Hello World"  
    def __init__(self):
        self._initial_attrs = set(vars(self).keys())
        self.init()
