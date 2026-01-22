import asyncio
import sys
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


async def main() -> None:
    server = StdioServerParameters(
        command=sys.executable,
        args=["-c", "from officemcp import main; main()"],
        env={"PYTHONUNBUFFERED": "1"},
    )

    async with stdio_client(server) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            tools = await session.list_tools()
            tool_names = [t.name for t in tools.tools]
            print(f"Tool count: {len(tool_names)}")

            expected = [
                "AvailableApps",
                "RunningApps",
                "IsAppAvailable",
                "DownloadImage",
                "RootFolder",
                "Visible",
                "Launch",
                "ScreenShot",
                "IsFileExists",
                "RunPython",
                "Quit",
                "Speak",
                "Beep",
                "Demonstrate",
            ]
            missing = [name for name in expected if name not in tool_names]
            if missing:
                raise SystemExit(f"Missing tools: {missing}")

            root_res = await session.call_tool("RootFolder", {})
            root = root_res.content[0].text if root_res.content else ""
            print(f"RootFolder: {root}")

            # ScreenShot end-to-end
            shot_name = "_smoke_screenshot.png"
            shot_res = await session.call_tool("ScreenShot", {"save_path": shot_name})
            shot_path = shot_res.content[0].text if shot_res.content else ""
            print(f"ScreenShot returned: {shot_path}")
            if shot_path and not Path(shot_path).exists():
                raise SystemExit(f"Screenshot file not found: {shot_path}")

            exists_res = await session.call_tool(
                "IsFileExists", {"sub_file_path": shot_name}
            )
            exists_text = exists_res.content[0].text if exists_res.content else ""
            print(f"IsFileExists('{shot_name}') returned: {exists_text}")

            beep_res = await session.call_tool(
                "Beep", {"frequency": 440, "duration": 150}
            )
            beep_text = beep_res.content[0].text if beep_res.content else ""
            print(f"Beep returned: {beep_text}")

            # DownloadImage (best-effort; may fail due to network restrictions)
            dl_name = "_smoke_download.ico"
            dl_res = await session.call_tool(
                "DownloadImage",
                {"url": "https://www.bing.com/favicon.ico", "save_path": dl_name},
            )
            dl_path = dl_res.content[0].text if dl_res.content else ""
            print(f"DownloadImage returned: {dl_path}")
            if dl_path and Path(dl_path).exists():
                print("DownloadImage file exists: OK")
            else:
                print("DownloadImage file exists: SKIP/FAIL (network or save error)")

    # Cleanup
    try:
        if root:
            for rel in [shot_name, dl_name]:
                p = Path(root) / rel
                if p.exists():
                    p.unlink()
                    print(f"Cleaned up: {p}")
    except Exception as e:
        print(f"Cleanup warning: {e}", file=sys.stderr)


if __name__ == "__main__":
    asyncio.run(main())
