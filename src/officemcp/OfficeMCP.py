# coding=utf-8
import sys
import time
from fastmcp import FastMCP
from fastmcp.resources import FileResource, TextResource, DirectoryResource
from officemcp.Officer import TheOfficer
mcp = FastMCP("OfficeMCP")
mcp.isrunnning = False
mcp.Officer = TheOfficer()
Officer = mcp.Officer
import inspect
#region automatically generate methods for wrighter's methods with @mcp.tool()        
    # class MCPclient:
    #     async def toolfunction(self,name:str, visible:bool=True) -> MCPclient_Call_Result:
    #         respose = await self.client.call_tool("AvailabeApps", {name:name, visible:visible})
    #         #respose is like [TextContent(type='text', text='[\n  "Word",\n  "Excel"n]', annotations=None)]#
    #         return MCPclient_Call_Result(respose)
    #     # async with MCPclient(SSETransport(url="http://localhost:8000/sse")) as client:
    #     #     result = await client.AvailableApps()
    #     #     if result:
    #     #         print(f"Available apps: {result.output}")
    #     #     else:
    #     #         print(f"Error: {result.error}")
# endregion

# region Office Apps ------------------------------

@mcp.tool()
def AvailableApps() -> list:
    """Get Microsoft Office applications availability. """
    return Officer.AvailableApps()

@mcp.tool()
def RunningApps() -> list:
    """Get Microsoft Office applications availability. """
    return Officer.RunningApps()

@mcp.tool()
def IsAppAvailable(app_name: str = "Word") -> bool:
    """ Check if the specified application is installed."""
    return Officer.IsAppAvailable(app_name)


@mcp.tool()
def DownloadImage(url: str='https://www.bing.com/favicon.ico', save_path: str='favicon.ico') -> str:
    """ Download an image from the given URL and save it to the specified path."""
    Officer.Print(f'Tool.DownloadImage....{url}  to  {save_path}')
    path =  Officer.DownloadImage(url, save_path)
    print('    path: {path}')

@mcp.tool()
def RootFolder() -> str:
    """ return the default folder for this OfficeMCP server."""
    return Officer.RootFolder


@mcp.tool()
def Visible(app_name: str="Word", visible: bool = True) -> bool:
    """ Check if the microsoft excel application is visible."""
    return Officer.Visible(app_name,visible)

@mcp.tool()
def Launch(app_name: str ="Word", visilbe: bool = True)->bool:
    """ Launch an new microsoft excel application or use the existed one."""
    Officer.Print('Tool.Launch....')
    try:
        app = Officer.Application(app_name)
        app.Visible = visilbe
        Officer.Print('    Launched:')
        return True
    except Exception as e:
        Officer.Print(f'    failed{e}:')
        return False

@mcp.tool()
def ScreenShot(save_path: str = None) -> str:
    """ Launch an new microsoft excel application or use the existed one."""
    Officer.Print('Tool.ScreenShot....')
    try:
        path  = Officer.ScreenShot(save_path)
        Officer.Print(f'   saved to {path}: ')
        return  path
    except Exception as e:
        Officer.Print(f'   failed {e}: ')
        return ""

@mcp.resource("resource://README.md")
def ReadME() -> TextResource:
    """ Read the readme file."""
    return TextResource(Officer.FilePath("README.md"))
# @mcp.resource("files://{file_name}")
# def DownloadFile(file_name: str):
#     return Officer.DownloadFile(file_name)

# @mcp.tool()
# def RootFolder() -> str:
#     """ 
#     The 
#     Get the root folder for this OfficeMCP server.
#     Don't create and use any folder or file outside this RootFolder all the time
#     Even you call tool RunPython(...) or call any type of office document save as\save
#     """
#     return Officer.RoodFolder()

@mcp.tool()
def IsFileExists(sub_file_path: str) -> bool:
    return Officer.IsFileExists(sub_file_path)

@mcp.tool()
def RunPython(code: str = "\nprint(f'hello world from {data}')\noutput=data\n", data:str = "OfficeMCP") -> dict:
    """ Run Python codes to control office applications with data.        
        Example:
            RunPython('print(data)','This is data content') #'This is data content' printed
            excel = Officer.Excel
            book = excel.Workbooks.Add() # add a new workbook
            sheet = excel.ActiveSheet
            excel.Visible = True
            sheet.Cells(1, 1).Value = "Hello, World From OfficeMCP!"
            sheet.Cells(3, 1).Value = "The data passed in:"
            sheet.Cells(4, 1).Value = data
            sheet.Cells(5, 1).Value = "output:"
            output = "Demo succesed"
            sheet.Cells(6, 1).Value = output     
    """
    print('Tool.RunPython')
    try:
        namespace = {
            'data': data,
            'Officer': mcp.Officer,
            'output': 'Execution completed',
            '__builtins__': __builtins__  # 确保内置函数可用
        }
        exec(code, namespace)
        return {'success': True, 'output': namespace['output']}
    except Exception as e:
        return {'success': False, 'error': str(e)}


@mcp.tool()
def Quit(app_name: str="Word",force:bool=False)->bool:
    """ Quit the microsoft excel application."""
    print('Tool.Quit:')
    return Officer.Quit(app_name,force)

@mcp.tool()
def Speak(text: str = "I'm office mcp server , how are you", volume: int = 80, rate: int = 0)->bool:
    """ Speak the text. volume range is 0-100, rate range is -10 to 10."""
    print('Tool.Speak:')
    return Officer.Speak(text, volume, rate)

@mcp.tool()
def Beep(frequency:int=500,duration:int=500)->bool:
    """ Beep the computer. frequency range is 37 to 32767, duration range is 0 to 65535."""
    print('Tool.Beep:')
    return Officer.Beep(frequency, frequency)
    
@mcp.tool()
def Demonstrate()->dict:
    """ Demonstrate for you to see some functions in this OfficeMCP server."""
    print('Tool.Demonstrate:')
    try:
        output = Officer.Demonstrate()    
        return {"success": True, "output": output}
    except Exception as e:
        print(e)
        return {"success": False, "error": str(e), "output": output}

@mcp.resource("resource://Instructions")
def Instructions() -> str:
    return """
    There're some base tools for you to control Microsoft applications.
    Specially you can use tool RunPython to run python codes to control Microsoft applications.
    There're an object called "Officer" as global, it have properties (Officer.Excel, Officer.Word, Officer.Outlook etc.) representing the Microsoft applications.
    """

# endregion

# region Browser stuff sync -----------
    # from playwright.sync_api import sync_playwright
    # wrighterSync = sync_playwright().start()
    # browser = wrighterSync.chromium.launch(headless=False)
    # @mcp.tool()
    # def NavigateSync(url: str = "https://www.bing.com")->bool:
    #     page = browser.new_page()
    #     page.goto('http://playwright.dev')
    #     page.screenshot(path=f'example-{wrighterSync}.png')
    #     return True          
# endregion ---------------------

configure_template="""
Put this in your mcp server config file or as the ide request:
{{
  \"mcpServers\": {{
    \"playwright\": {{
      \"url\": \"http://{host}:{port}/sse\"
    }}
  }}
}}
"""

def RunOfficeMCP() -> None:
    r"""OfficeMCP server entry point with command line arguments support.
    Usage examples:
    1. OfficeMCP (stdio mode by default)
    2. OfficeMCP sse --port 8080 --host 127.0.0.1 --folder d:\@OfficeMCP (alternative syntax)
    """
    import os
    args = sys.argv[1:]
    transport = "stdio"
    thePort = 8888
    theHost = "127.0.0.1"
    theFolder = "D:\\@OfficeMCP"

    # 更健壮的参数解析
    i = 1
    if args:
        transport = args[0].lower()
        while i < len(args):
            arg = args[i]
            if arg == '--port' and i+1 < len(args):
                try:
                    thePort = int(args[i+1])
                except Exception:
                    pass
                i += 2
            elif arg == '--host' and i+1 < len(args):
                theHost = args[i+1]
                i += 2
            elif arg == '--folder' and i+1 < len(args):
                theFolder = args[i+1]
                i += 2
            elif arg.isdigit():
                thePort = int(arg)
                i += 1
            elif not arg.startswith('--') and theHost == "127.0.0.1":
                # 只有没有明确指定 --host 时才用
                theHost = arg
                i += 1
            else:
                i += 1

    # 校验文件夹路径
    if not theFolder or not isinstance(theFolder, str) or not os.path.isabs(theFolder):
        theFolder = "D:\\@OfficeMCP"
    if not os.path.exists(theFolder):
        try:
            os.makedirs(theFolder)
        except Exception:
            print(f"Warning: Could not create folder {theFolder}, fallback to D:\\@OfficeMCP")
            theFolder = "D:\\@OfficeMCP"
            if not os.path.exists(theFolder):
                os.makedirs(theFolder)
    mcp.Officer._default_folder = theFolder

    try:
        if transport == "stdio":
            print(f"OfficeMCP running in stdio mode, root folder: {theFolder}")
            mcp.run("stdio")
        else:
            print(configure_template.format(host=theHost, port=thePort))
            print(f"OfficeMCP running in SSE mode, host: {theHost}, port: {thePort}, root folder: {theFolder}")
            mcp.run(transport="sse", host=theHost, port=thePort)
    except ValueError as e:
        print(f"OfficeMCP Error parsing arguments: {e}")
    except Exception as e:
        print(f"OfficeMCP Server startup failed: {e}")
