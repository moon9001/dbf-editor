# DBF Editor

一个基于 Web 的 DBF 文件编辑器，支持查看、编辑、搜索、筛选、导入导出、转 Excel，固定端口局域网访问。

## 功能

| 功能 | 说明 |
|------|------|
| 📋 查看 | 分页表格，支持排序 + 搜索 + 字段筛选 |
| ✏️ 编辑 | 双击单元格直接编辑 / 整行编辑弹窗 |
| ➕ 增删 | 添加新行、单行删除、批量删除 |
| 📊 导出 Excel | 一键导出 .xlsx，带表头样式、冻结首行、自动筛选 |
| 💾 导出 DBF | 下载原始 DBF 文件，格式编码字段完全不变 |
| 📥 从 Excel 导入 | 从 Excel 导入数据，支持新建或追加 |
| 🔍 字段筛选 | 自动检测编码字段，生成下拉筛选器（性别/考试方式等） |
| 🔑 只读字段 | DWDM/DWMC 等单值字段自动保护，防止误改 |
| 📂 文件管理 | 管理面板列出所有上传文件，支持一键删除 |
| 🌐 局域网访问 | 固定端口 5000，启动后显示局域网 IP |

## 快速开始

### 环境要求

- Python 3.8+
- Windows / Linux / macOS

### 安装

```bash
# 1. 克隆仓库
git clone https://github.com/YOUR_USERNAME/dbf-editor.git
cd dbf-editor

# 2. 创建虚拟环境
python -m venv venv

# 3. 安装依赖
# Windows:
venv\Scripts\pip install flask dbf openpyxl
# Linux/macOS:
venv/bin/pip install flask dbf openpyxl
```

### 启动

**Windows** — 双击 `start_dbf_editor.bat`

**手动启动**：
```bash
venv\Scripts\python.exe server.py   # Windows
venv/bin/python server.py            # Linux/macOS
```

浏览器访问 http://127.0.0.1:5000

### 添加 DBF 文件

将你的 `.dbf` 文件放到 `dbf_editor/` 同级目录下（如 `80106BBK.dbf`），程序会自动检测。或通过网页"📤 上传"按钮上传。

## 项目结构

```
├── start_dbf_editor.bat      ← Windows 一键启动
├── DBF编辑器使用说明.md        ← 中文详细说明
├── server.py                 ← Flask 后端
├── templates/
│   └── index.html            ← 前端界面
├── requirements.txt          ← Python 依赖
├── README.md
└── .gitignore
```

## 技术栈

- **后端**: Python Flask + dbf 库 + openpyxl
- **前端**: 原生 HTML/CSS/JS（无框架依赖）
- **编码**: GBK (cp936) 中文支持

## License

MIT
