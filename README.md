[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/zh19980811-easy-mcp-autocad-badge.png)](https://mseep.ai/app/zh19980811-easy-mcp-autocad)


# AutoCAD MCP 服务器 / AutoCAD MCP Server

> ⚠️ 本项目目前维护精力有限，欢迎有兴趣的开发者参与协作！  
> ⚠️ Currently this project is **not actively maintained** due to time constraints. I would be **very happy to collaborate** with anyone interested in co-maintaining or extending it.

基于 **Model Context Protocol (MCP)** 的 AutoCAD 集成服务器，允许通过 **Claude** 等大型语言模型 (LLM) 与 AutoCAD 进行自然语言交互。  
An AutoCAD integration server based on **Model Context Protocol (MCP)**, enabling natural language interaction with AutoCAD via large language models like **Claude**.

> 🔗 项目在 MseeP.ai 展示 / Referenced on MseeP.ai:  
> https://mseep.ai/app/zh19980811-easy-mcp-autocad

🎬 **演示视频 / Demo Video**:  
[![AutoCAD MCP 演示视频](https://img.youtube.com/vi/-I6CTc3Xaek/0.jpg)](https://www.youtube.com/watch?v=-I6CTc3Xaek)

---

## ✨ 功能特点 / Features

- 🗣️ 自然语言控制 AutoCAD 图纸 / Natural language control of AutoCAD drawings  
- ✏️ 基础绘图功能（线条、圆）/ Basic drawing tools (line, circle)  
- 📚 图层管理 / Layer management  
- 🧠 自动生成 PMC 控制图 / Auto-generate PMC control diagrams  
- 🔍 图纸元素分析 / Drawing element analysis  
- 🔦 文本高亮匹配 / Highlight specific text patterns  
- 💾 内嵌 SQLite 数据库存储 CAD 元素 / Integrated SQLite storage for CAD elements  

---

## 🖥️ 系统要求 / System Requirements

- Python 3.10+  
- AutoCAD 2018+ (需支持 COM 接口 / with COM interface support)  
- Windows 操作系统 / Windows OS  

---

## ⚙️ 安装步骤 / Installation

### 1. 克隆仓库 / Clone the repository

```bash
git clone https://github.com/yourusername/autocad-mcp-server.git
cd autocad-mcp-server

| `export_to_database` | 将 CAD 元素信息存入 SQLite |

## 2. 创建虚拟环境 / Create virtual environment

**Windows:**

```bash
python -m venv .venv
.venv\Scripts\activate
```

**macOS / Linux:**

```bash
python -m venv .venv
source .venv/bin/activate
```

---

## 3. 安装依赖 / Install dependencies

```bash
pip install -r requirements.txt
```

---

## 4. （可选）构建可执行文件 / (Optional) Build as executable

```bash
pyinstaller --onefile server.py
```

---

## 🚀 使用方法 / How to Use

### 独立运行服务器 / Run server independently

```bash
python server.py
```

---

### 集成 Claude Desktop / Integrate with Claude Desktop

编辑配置文件：

**Windows 路径 / Config path on Windows:**

```shell
%APPDATA%\Claude\claude_desktop_config.json
```

**示例配置 / Example config:**

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

---

## 🧰 工具 API / Available API Tools

| 功能 / Function         | 描述 / Description                       |
|------------------------|------------------------------------------|
| `create_new_drawing`   | 创建新的图纸 / Create a new drawing      |
| `draw_line`            | 画直线 / Draw a line                     |
| `draw_circle`          | 画圆 / Draw a circle                     |
| `set_layer`            | 设置当前图层 / Set current drawing layer |
| `highlight_text`       | 高亮显示匹配文本 / Highlight matching text |
| `scan_elements`        | 扫描并解析图纸元素 / Scan and parse drawing elements |
| `export_to_database`   | 导出 CAD 元素到数据库 / Export CAD data to SQLite |

---

## 🙋‍♂️ 维护状态 / Maintenance Notice

⚠️ 当前我正忙于其他项目，维护精力有限。欢迎 Fork 项目或提交 PR，一起完善 AutoCAD 智能交互生态！  
⚠️ I'm currently busy and not able to actively maintain this repo. PRs and collaborators are welcome!

📬 联系我 / Contact: 1062723732@qq.com 

---

Made with ❤️ for open-source learning.
