import asyncio
import sys

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

    print(f"Tool count: {len(tools.tools)}")
    for tool in tools.tools:
        print(f"- {tool.name}")


if __name__ == "__main__":
    asyncio.run(main())
