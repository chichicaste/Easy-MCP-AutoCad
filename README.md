# AutoCAD MCP服务器

基于Model Context Protocol (MCP)的AutoCAD集成服务器，允许通过Claude等大型语言模型(LLM)与AutoCAD进行自然语言交互。

## 功能特点

- 通过自然语言控制AutoCAD创建和修改图纸
- 绘制基本图形(线条、圆等)
- 创建并管理图层
- 自动生成PMC控制图等专业图纸
- 扫描并分析现有图纸中的元素
- 查询和高亮显示特定文本模式(如PMC-3M)
- 集成SQLite数据库，存储和查询CAD元素信息

## 系统要求

- Python 3.10或更高版本
- AutoCAD 2018或更高版本
- Windows操作系统(与AutoCAD COM接口兼容)

## 安装

1. 克隆仓库
git clone https://github.com/yourusername/autocad-mcp-server.git
cd autocad-mcp-server
Copy
2. 创建并激活虚拟环境
python -m venv .venv
.venv\Scripts\activate  # Windows
source .venv/bin/activate  # macOS/Linux
Copy
3. 安装依赖
pip install -r requirements.txt
Copy
4. 构建可执行文件(可选)
pyinstaller --onefile server.py
Copy
## 使用方法

### 作为独立服务器运行
python server.py
Copy
### 与Claude Desktop集成

编辑Claude Desktop配置文件(`~/Library/Application Support/Claude/claude_desktop_config.json`或`%APPDATA%\Claude\claude_desktop_config.json`)：

```json
{
  "mcpServers": {
    "autocad-db-server": {
      "command": "path/to/autocad_mcp_server.exe",
      "args": []
    }
  }
}
可用工具

create_new_drawing - 创建新的AutoCAD图纸
draw_line - 绘制直线
draw_circle - 绘制圆形
scan_all_entities - 扫描当前图纸中的所有实体并保存到数据库
count_text_patterns - 统计文本模式(如PMC-3M)的出现次数
highlight_text_matches - 高亮显示包含特定模式的文本
create_pmc_control_diagram - 创建PMC控制设备图
execute_autocad_script - 执行AutoCAD脚本
