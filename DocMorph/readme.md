
## 📄 Document Converter 文档格式转换工具

一个支持 **多种文档格式互转** 的 Python 工具，包含命令行与 GUI 界面，支持 PDF、Word、Markdown、Excel、HTML 等常见格式之间的转换。

> **支持功能**：单文件转换、多文件批量转换、文件夹一键转换、进度条显示、转换失败提示

------

## 🔧 支持的格式转换

| 源格式  | 可转换目标格式         |
| ------- | ---------------------- |
| `.pdf`  | `.docx`, `.txt`        |
| `.docx` | `.pdf`, `.md`, `.html` |
| `.xlsx` | `.csv`                 |
| `.md`   | `.docx`, `.html`       |
| `.html` | `.docx`, `.pdf`, `.md` |

------

## 🖥️ 使用方式

### ✅ 1. 命令行使用（converter.py）

```bash
# 单文件转换（自动识别格式）
python converter.py --input ./input/sample.docx --output ./output/sample.pdf

# 指定源格式和目标格式（适用于批量）
python converter.py --input ./input --output ./output --from docx --to pdf --batch
```

默认路径：

- 输入目录：`./input`
- 输出目录：`./output`

------

### ✅ 2. 图形界面使用（converter_gui.py）

```bash
python converter_gui.py
```

#### 功能特性：

- ✅ 单文件转换（点击选择）
- ✅ 多文件批量转换（可多选文件）
- ✅ 整个文件夹批量转换（自动遍历指定目录）
- ✅ 一键转换 ./input 目录下所有支持的文件
- ✅ 显示转换进度条与失败文件名

------

## 🧱 项目结构

```
.
├── converter.py           # 核心转换逻辑
├── converter_gui.py       # GUI 图形界面
├── input/                 # 默认输入文件夹
├── output/                # 默认输出文件夹
├── bin/                   # 可选：打包 pandoc.exe 用
└── README.md              # 项目说明文件
```

------

## 🧩 安装依赖

```bash
pip install -r requirements.txt
```

**推荐依赖列表（requirements.txt）：**

```
pdf2docx
docx2pdf
pypandoc
pandas
PyPDF2
tk
```

------

## ⚠️ 关于 Pandoc 依赖

### ❗重要：`pypandoc` 只是一个调用器，真正依赖的是系统中的 `pandoc` 工具。

你有以下三种方式解决发布时的兼容问题：

------

### ✅ 方式一：随程序打包 `pandoc.exe`（推荐）

1. 从官网下载安装 Pandoc portable：
    https://github.com/jgm/pandoc/releases
2. 解压后将 `pandoc.exe` 放入你的项目目录下（如 `./bin/pandoc.exe`）
3. 在代码中添加：

```python
import pypandoc
import os
pypandoc.pandoc_path = os.path.abspath("./bin/pandoc.exe")
```

------

### ✅ 方式二：首次运行时自动下载 Pandoc（需联网）

```python
import pypandoc
pypandoc.download_pandoc()
```

📦 缺点：用户首次运行需联网，且下载较慢（约 80MB）

------

### ✅ 方式三：强依赖系统环境

你也可以要求用户手动安装 Pandoc 并配置 PATH，但不推荐这样做发布。

------

## 📦 打包发布建议

如你使用 `PyInstaller` 打包成 `.exe`：

```bash
pyinstaller --noconfirm --onefile converter_gui.py
```

请确保 `pandoc.exe` 一并打入包内，可在 `.spec` 文件中加入：

```python
datas = [('./bin/pandoc.exe', 'bin')]
```

------


📮 反馈与贡献

欢迎在提交反馈与建议，欢迎 star 🌟、fork 🍴 或参与改进。
