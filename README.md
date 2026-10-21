<div align="center">

<!-- omit in toc -->
# OfficeMCP v1.0.5
<strong>>The most seeable and free way to control Microsof applications by AI model.</strong>

[![PyPI - Version](https://img.shields.io/pypi/v/officemcp.svg)](https://pypi.org/project/officemcp)

</div>

# OfficeMCP

OfficeMCP server is designed for AI to automate Microsoft Office Applications (Word, Excel, PowerPoint, Access, OneNote, Visio, Project, WPS.word, Wps.powerpoint, wps.excel etc.) by COM interface in Windows OS. Not working on Linux/MacOS.

# Warns

Please keep it in mind, as OfficeMCP not limit the usage of python. epeciall there's a tool RunPython(...) to execute python codes created by Ai model. But it is also the most wonderfull parts of OfficeMCP. we can't guarantee that your AI model will not do something bad to your computer. we don't take any responsibility.

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
- The sse server is a separate process. Nothing starts it for you, so it has to be running before a client connects, and it stops when the shell that runs it is closed. A client that reports a connection error on every start is usually pointing at this server while it is not running.
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

### 3. Use OfficeMCP with OpenCode

OpenCode reads `~/.config/opencode/opencode.jsonc` and nests servers under `mcp.servers`:

```json
{
    "$schema": "https://opencode.ai/config.json",
    "mcp": {
        "servers": {
            "OfficeMCP": {
                "type": "local",
                "command": [
                    "uvx",
                    "officemcp"
                ]
            }
        }
    }
}
```

Two things are worth knowing here:

- OpenCode expects `mcp.servers`. A config that puts the server directly under `mcp` uses the older layout, which is deprecated but still loads.
- A project can override a server by defining the same name in an `opencode.jsonc` in the project root. The project entry replaces the global one rather than merging with it, so it has to repeat the type and command.

On Windows, `uv` and `uvx` install to `%USERPROFILE%\.local\bin`. That folder is on the user PATH, so `uvx` resolves by name, but you can also point the command at `C:\\Users\\<you>\\.local\\bin\\uvx.exe` if the client does not inherit your PATH.

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
git clone https://github.com/OfficeMCP/OfficeMCP
```
