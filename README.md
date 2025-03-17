AutoCAD MCP 服务器
基于 Model Context Protocol (MCP) 的 AutoCAD 集成服务器，允许通过 Claude 等大型语言模型 (LLM) 与 AutoCAD 进行自然语言交互。

功能特点
自然语言交互：通过自然语言控制 AutoCAD 创建和修改图纸
基础绘图：支持绘制基本图形（线条、圆等）
图层管理：创建、修改和删除图层
专业图纸生成：自动生成 PMC 控制图 等专业图纸
图纸分析：扫描并解析现有图纸中的元素
文本模式查询：查询并高亮显示特定文本模式（如 PMC-3M）
数据库集成：内置 SQLite 数据库，支持 CAD 元素的存储和查询
系统要求
Python 3.10 或更高版本
AutoCAD 2018 或更高版本（需支持 COM 接口）
Windows 操作系统
安装
1. 克隆仓库
sh
复制
编辑
git clone https://github.com/yourusername/autocad-mcp-server.git
cd autocad-mcp-server
2. 创建并激活虚拟环境
Windows：

sh
复制
编辑
python -m venv .venv
.venv\Scripts\activate
macOS / Linux：

sh
复制
编辑
python -m venv .venv
source .venv/bin/activate
3. 安装依赖
sh
复制
编辑
pip install -r requirements.txt
4. （可选）构建可执行文件
sh
复制
编辑
pyinstaller --onefile server.py
使用方法
作为独立服务器运行
sh
复制
编辑
python server.py
与 Claude Desktop 集成
编辑 Claude Desktop 配置文件（路径如下）：

Windows: %APPDATA%\Claude\claude_desktop_config.json
macOS: ~/Library/Application Support/Claude/claude_desktop_config.json
示例配置：

json
复制
编辑
{
  "mcpServers": {
    "autocad-mcp-server": {
      "command": "path/to/autocad_mcp_server.exe",
      "args": []
    }
  }
}
