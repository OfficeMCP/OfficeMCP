# OfficeMCP

OfficeMCP server is designed for AI to automate Microsoft Office Applications (Word, Excel, PowerPoint, Access, OneNote, Visio, Project, WPS.word, Wps.powerpoint, wps.excel etc.) by COM interface in Windows OS. Not working on Linux/MacOS.

# Warns

you should this mcp carefully, as it not limit the usage of python. epeciall there's a tool RunPython(...) to execute Ai model's thought, but it is also the wonderfull parts of OfficeMCP. we can't guarantee that your AI model will not do something bad to your computer. we don't take any responsibility.

# System Requirements

1. Windows system

2. python 3.1 or above installed

3. uv installed
    open an shell window and run command    
    >pip install uv



## How to install OfficeMCP
There are two ways or two modes to install OfficeMCP (They also can be used in the same time):

### 1. Use OfficeMCP as stdio server: 
- One OfficeMCP server for one mcp client mode
- Put following setting to MCP.json file for vscode or some proper place for other AI IDE:

```json
{
    "mcpServers": {
        "OfficeMCP": {
            "type": "stdio",
            "command": "uvx",
            "args": [
                "officemcp"
            ]
        }
    }
}
```

### 2. Use OfficeMCP as sse server: 
- One OfficeMCP server for multi mcp client mode
- You can change port and host as you like
- This is recommended way to use OfficeMCP server.
#### step 1:  
**Run one command in shell or power shell:**
>uvx officemcp sse

the Mcp server url will be:  "http://127.0.0.1:8888/sse"  or  "http://127.0.0.1:8888/sse" 
the default work folder is D:\@officemcp


##### or something like below
>uvx officemcp sse --port 7777 --host 127.0.0.8 --folder D:\myfolder

"url" will be : "http://127.0.0.8:7777/sse"
#### step 2: 
**Put following setting to MCP.json file for vscode or some proper place for other AI IDE:**

```json
{
    "servers": {
        "OfficeMCP": {
            "url": "http://127.0.0.1:8888/sse"
        }
    }
}
```

####  or

```json
{
    "servers": {
        "OfficeMCP": {
            "url": "http://{your_host}:{your_port}/sse"
        }
    }
}

```

## Usage
On AI IDE, you can ask AI model to control Office Applications by OfficeMCP server:
- You ask AI model to open a new Office Application.
    AI model will send a request to OfficeMCP server, and OfficeMCP server will open a new Office Application.

- You ask AI model to do whatever you want to do in the current Office Application.
    AI model will analyze your request, and call OfficeMCP server's tool to accomplish your request.

## Tools Reference
Tools:
- AvailableApps(): check if Microsoft Office applications are installed on your computer.

- RunningApps(): get a list of currently running Office applications.

- IsAppAvailable(...): check if a specific Office application is installed.

- Launch(...): launch a new Office application and set its visibility.

- Visible(...): set the specified Office application's visibility to True or False.

- Quit(...): quit the specified Office application.

- Demonstrate(): run a demonstration of OfficeMCP automation features.

- Speak(...): speak a string you passed in.

- Beep(...): play a beep sound.

- DefaultFolder(...): return the OfficeMCP root work folder default is ("D:\OfficeMCP")

- IsFileExists(sub_path): check if a file exists in the OfficeMCP root folder.

- DownloadImage(...): download an image from a given URL and save it to the specified path.

- RunPython(codes,data): run python code in the OfficeMCP server context.
    - This is the most powerful tool in OfficeMCP server. AI can use this tool to do anything supported by the server, including automating Office applications.
    - There is an object "Officer" that can be used in the python code, e.g. `Officer.Excel` holds the current Excel com Application, and more are Officer.Word, Officer.Powerpoint, Office.Visio, Officer.Access, Officer.OneNote, Officer.Visio, Officer.Project. Office.Kwps for WPS word, Office.Ket for WPS excell, Office.Kwpp for WPS powerpoint.
    - There is an object "output" as RunPython(...) return that can be used in the python code, to put your own return result in to output, like output="run python sccessed", then RunPython return "run python sccessed" to AI model.
    - You can use Officer.Visio to create a new Visio document, and then use Officer.Visio.ActivePage to get the active page, and use Officer.Visio.ActivePage.DrawRectangle(...) to draw a rectangle on the page.
    - You can use Officer.Excel to create a new Excel document, and then use Officer.Excel.ActiveSheet to get the active sheet, and use Officer.Excel.ActiveSheet.Cells(...) to get the cell, and use Officer.Excel.ActiveSheet.Cells(...).Value = "hello" to set the cell value.
    - You use codes to control them by running the codes with RunPython tool.

- More tools will be added in the future.


## Development
```bash
git clone https://github.com/officemcp/officemcp
```
