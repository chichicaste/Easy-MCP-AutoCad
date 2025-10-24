[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/zh19980811-easy-mcp-autocad-badge.png)](https://mseep.ai/app/zh19980811-easy-mcp-autocad)

# AutoCAD MCP Server

WARNING: This project is currently not actively maintained due to limited time. Contributions and collaborators are welcome.

## Overview
An AutoCAD integration server implementing the Model Context Protocol (MCP). It enables natural language interaction with AutoCAD through large language models (LLMs) such as Claude.

Referenced on MseeP.ai
https://mseep.ai/app/zh19980811-easy-mcp-autocad

Demo
Demo video:
[AutoCAD MCP Demo](https://www.youtube.com/watch?v=-I6CTc3Xaek)

---

## Features

- Natural-language control of AutoCAD drawings
- Basic drawing tools (line, circle)
- Layer management (create, modify, delete)
- Auto-generate PMC control diagrams and other specialized drawings
- Scan and analyze existing drawings to extract elements
- Highlight and search for text patterns (e.g., PMC-3M)
- Integrated SQLite database to store and query CAD elements

## System Requirements

- Python 3.10+
- AutoCAD 2018+ (must support COM automation)
- Windows OS

## Installation

1. Clone the repository
```bash
git clone https://github.com/yourusername/autocad-mcp-server.git
cd autocad-mcp-server
```

2. Create and activate a virtual environment

Windows:
```bash
python -m venv .venv
.venv\Scripts\activate
```

macOS / Linux:
```bash
python -m venv .venv
source .venv/bin/activate
```

3. Install dependencies
```bash
pip install -r requirements.txt
```

4. (Optional) Build a standalone executable
```bash
pyinstaller --onefile server.py
```

## Usage

Run the server as a standalone process:
```bash
python server.py
```

## Integrate with Claude Desktop

Edit the Claude Desktop configuration file (example paths):
- Windows: %APPDATA%\Claude\claude_desktop_config.json
- macOS: ~/Library/Application Support/Claude/claude_desktop_config.json

Example Claude config:
```json
{
  "mcpServers": {
    "autocad-mcp-server": {
      "command": "path/to/autocad_mcp_server.exe",
      "args": []
    }
  }
}
```

## Available API Tools

| Function              | Description                                 |
|-----------------------|---------------------------------------------|
| `create_new_drawing`  | Create a new AutoCAD drawing                |
| `draw_line`           | Draw a straight line                        |
| `draw_circle`         | Draw a circle                               |
| `set_layer`           | Set the current drawing layer               |
| `highlight_text`      | Highlight matching text in the drawing      |
| `scan_elements`       | Scan and parse drawing elements             |
| `export_to_database`  | Export CAD element data to the bundled SQLite DB |

## Maintenance

This repository is not actively maintained at the moment. Pull requests and collaborators are welcome.

Contact: 1062723732@qq.com

---

Made with ❤️ for open-source learning.
