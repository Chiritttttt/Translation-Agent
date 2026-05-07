# Translation Agent 桌面版 — 打包与部署指南

## 目录

- [1. 环境要求](#1-环境要求)
- [2. 快速开始](#2-快速开始)
- [3. PyCharm 配置](#3-pycharm-配置)
- [4. 构建模式详解](#4-构建模式详解)
- [5. 构建脚本使用](#5-构建脚本使用)
- [6. 配置文件说明](#6-配置文件说明)
- [7. 常见问题排查](#7-常见问题排查)
- [8. 分发与部署](#8-分发与部署)

---

## 1. 环境要求

### 必需

| 项目 | 要求 |
|------|------|
| Python | 3.9 ~ 3.12（推荐 3.11） |
| pip | 最新版 |
| 操作系统 | Windows 10+ / macOS 12+ / Ubuntu 20.04+ |
| 磁盘空间 | 构建时至少 3 GB 可用 |

### 可选（PDF OCR 功能）

| 项目 | 说明 |
|------|------|
| poppler | `pdf2image` 依赖。Windows: `choco install poppler`；macOS: `brew install poppler`；Linux: `sudo apt install poppler-utils` |

### 安装 Python

**Windows:**
```powershell
# 推荐使用 winget
winget install Python.Python.3.11
```

**macOS:**
```bash
brew install python@3.11
```

**Ubuntu/Debian:**
```bash
sudo apt update
sudo apt install python3.11 python3.11-venv python3-pip
```

---

## 2. 快速开始

### 方式一：使用构建脚本（推荐）

**Windows:**
```powershell
cd translation-agent
# 构建两种版本
build.bat

# 或者只构建单文件版
build.bat onefile

# 只构建目录版
build.bat onedir

# 清理构建产物
build.bat clean
```

**macOS / Linux:**
```bash
cd translation-agent
chmod +x build.sh

# 构建两种版本
./build.sh

# 或者只构建单文件版
./build.sh onefile

# 只构建目录版
./build.sh onedir

# 清理构建产物（包括虚拟环境）
./build.sh clean
```

### 方式二：手动构建

```bash
# 1. 创建虚拟环境
python -m venv .venv-build
source .venv-build/bin/activate   # Linux/macOS
# 或 .venv-build\Scripts\activate  # Windows

# 2. 安装依赖
pip install -r requirements-desktop.txt

# 3. 构建（二选一）
pyinstaller TranslationAgent-onefile.spec --clean
pyinstaller TranslationAgent-onedir.spec --clean
```

### 构建产物位置

```
translation-agent/
├── dist/
│   ├── TranslationAgent          ← onefile：单个可执行文件
│   ├── TranslationAgent/         ← onedir：目录结构
│   │   ├── TranslationAgent      ← 主程序
│   │   └── _internal/            ← 依赖库
```

---

## 3. PyCharm 配置

### 3.1 配置运行环境

1. **File → Settings → Project → Python Interpreter**
2. 点击齿轮 → **Add Interpreter → Add Local Interpreter**
3. 选择 **Existing environment**
4. 路径指向 `.venv-build/bin/python`（或 Windows 下 `.venv-build\Scripts\python.exe`）

### 3.2 配置 PyInstaller 运行配置

1. **Run → Edit Configurations → + → Python**
2. 配置如下：

| 配置项 | 值 |
|--------|-----|
| Name | `Build Onefile` |
| Script path | PyInstaller 的启动脚本路径 |
| Parameters | `TranslationAgent-onefile.spec --clean --noconfirm` |
| Working directory | 项目 `translation-agent/` 目录 |
| Python interpreter | 项目虚拟环境 |

**或者用更简单的方式：** 在 PyCharm 的 Terminal 中直接运行构建命令。

### 3.3 常用 Terminal 命令

在 PyCharm 底部的 Terminal 标签中：

```bash
# 激活环境
source .venv-build/bin/activate

# 快速测试
python gui.py

# 构建并测试
pyinstaller TranslationAgent-onefile.spec --clean --noconfirm
./dist/TranslationAgent
```

---

## 4. 构建模式详解

### onefile（单文件模式）

| 项 | 说明 |
|----|------|
| 输出 | 单个 `TranslationAgent.exe`（或 macOS/Linux 可执行文件） |
| 大小 | 约 150-250 MB（含所有依赖） |
| 启动速度 | 首次启动较慢（需解压临时文件到 `%TEMP%`） |
| 优点 | 分发简单，只需一个文件 |
| 缺点 | 启动慢、防病毒软件可能误报、更新需下载整个文件 |
| 适用 | 个人使用、内部小范围分发 |

### onedir（目录模式）

| 项 | 说明 |
|----|------|
| 输出 | `TranslationAgent/` 目录（可执行文件 + `_internal/` 依赖目录） |
| 大小 | 总计约 200-300 MB |
| 启动速度 | 快（无需解压） |
| 优点 | 启动快、稳定可靠、易于调试 |
| 缺点 | 需要整个文件夹一起分发 |
| 适用 | 正式发布、企业部署、需要稳定的场景 |

### 推荐

> **开发/测试阶段**：直接运行 `python gui.py`
> **个人使用**：onefile 方便
> **正式发布**：onedir 更稳定

---

## 5. 构建脚本使用

### build.bat（Windows）

```
build.bat              # 同时构建 onefile 和 onedir
build.bat onefile      # 只构建单文件版
build.bat onedir       # 只构建目录版
build.bat clean        # 清理 build/ 和 dist/
```

脚本会自动：
1. 创建 `.venv-build` 虚拟环境
2. 安装 `requirements-desktop.txt` 中的依赖
3. 运行 PyInstaller
4. 复制 `.env.example` 到输出目录

### build.sh（macOS/Linux）

```bash
./build.sh              # 同时构建 onefile 和 onedir
./build.sh onefile      # 只构建单文件版
./build.sh onedir       # 只构建目录版
./build.sh clean        # 清理 build/、dist/ 和 .venv-build
```

---

## 6. 配置文件说明

### .env 文件

打包后的程序会从以下位置（按优先级）查找 `.env` 文件：

1. 可执行文件所在目录
2. 当前工作目录
3. 脚本所在目录（开发模式）

**首次使用：**

```bash
# 复制模板
cp .env.example .env

# 编辑配置
# Windows: notepad .env
# macOS/Linux: vim .env 或 nano .env
```

**必填项：**
```
OPENAI_API_KEY=sk-your-api-key-here
```

**可选项（使用 OpenAI 时无需配置）：**
```
OPENAI_BASE_URL=https://api.deepseek.com/v1
OPENAI_MODEL=deepseek-chat
```

### 支持的 API 服务商

| 服务商 | OPENAI_BASE_URL | 推荐模型 |
|--------|-----------------|----------|
| OpenAI | `https://api.openai.com/v1` | `gpt-4o` |
| DeepSeek | `https://api.deepseek.com/v1` | `deepseek-chat` |
| 智谱 AI | `https://open.bigmodel.cn/api/paas/v4` | `glm-4-flash` |
| 月之暗面 | `https://api.moonshot.cn/v1` | `moonshot-v1-8k` |
| 通义千问 | `https://dashscope.aliyuncs.com/compatible-mode/v1` | `qwen-plus` |
| 本地代理 | `http://localhost:8080/v1` | 视代理而定 |

---

## 7. 常见问题排查

### 7.1 构建失败

#### `ModuleNotFoundError: No module named 'tiktoken_ext.openai_public'`

这是最常见的问题。确保 `.spec` 文件的 `hiddenimports` 中包含：
```python
'tiktoken',
'tiktoken_ext',
'tiktoken_ext.openai_public',
```
并且使用 `collect_data_files('tiktoken')` 收集 BPE 数据文件。

#### `ImportError: DLL load failed`（Windows）

通常因为缺少 Visual C++ 运行时。安装 [Microsoft Visual C++ Redistributable](https://aka.ms/vs/17/release/vc_redist.x64.exe)。

#### PyQt6 相关错误

确保 PyQt6 完整安装：
```bash
pip uninstall PyQt6 PyQt6-Qt6 PyQt6-sip -y
pip install PyQt6
```

### 7.2 打包后运行失败

#### 启动闪退（无错误信息）

1. 将 `console=True` 写入 spec 文件重新构建
2. 或查看日志文件：`logs/translation_agent.log`

#### `.env 文件找不到`

将 `.env` 放在 **可执行文件同级目录**，不是 `_internal/` 目录。

#### `OPENAI_API_KEY 未配置`

在 `.env` 中设置，或在程序界面中配置。

#### 防病毒软件误报（Windows）

- **Windows Defender**：设置 → 隐私和安全 → Windows 安全中心 → 病毒和威胁防护 → 排除项 → 添加排除文件夹
- **其他杀毒软件**：将构建产物目录加入白名单
- **长期方案**：使用代码签名证书（需购买）

#### onefile 模式启动慢

这是正常行为。onefile 模式每次启动需将约 200MB 文件解压到临时目录。改用 onedir 模式可大幅提速。

#### PDF 导出中文乱码

程序会自动搜索系统 CJK 字体。如果仍显示乱码：
- **Windows**：确认已安装微软雅黑
- **Linux**：`sudo apt install fonts-noto-cjk`
- **macOS**：系统自带，无需额外安装

#### PDF OCR 功能不可用

需要安装 poppler：
```bash
# Windows (Chocolatey)
choco install poppler

# macOS
brew install poppler

# Ubuntu/Debian
sudo apt install poppler-utils
```

### 7.3 调试技巧

#### 开启控制台窗口

在 `.spec` 文件中设置：
```python
exe = EXE(
    ...
    console=True,    # 改为 True，显示控制台输出
)
```

#### 检查打包内容

```bash
# 列出 onefile 中包含的所有模块
pyinstaller --log-level DEBUG TranslationAgent-onefile.spec 2>&1 | grep "Adding"

# 检查 onedir 输出
ls -la dist/TranslationAgent/_internal/
```

#### 验证隐藏导入

```bash
# 进入虚拟环境
source .venv-build/bin/activate

# 测试模块是否可导入
python -c "import tiktoken_ext.openai_public; print('OK')"
python -c "import simplemma; print('OK')"
python -c "import fitz; print('OK')"
```

---

## 8. 分发与部署

### 8.1 Windows

**onefile 分发：**
1. 将 `dist/TranslationAgent.exe` 打包为 ZIP
2. 附带 `.env.example` 和说明文档
3. 用户解压后，复制 `.env.example` 为 `.env` 并填入 API Key

**onedir 分发：**
1. 将整个 `dist/TranslationAgent/` 目录打包为 ZIP
2. 附带 `.env.example`
3. 用户解压后，将 `.env` 放在 `TranslationAgent.exe` 同级目录

### 8.2 macOS

**注意**：macOS 要求对应用进行签名才能正常运行。

```bash
# 临时方案：移除隔离属性
xattr -cr dist/TranslationAgent

# 长期方案：代码签名（需要 Apple Developer 账号）
codesign --deep --force --verify --verbose --sign "Developer ID Application: Your Name" dist/TranslationAgent

# 创建 .app bundle（手动）
mkdir -p TranslationAgent.app/Contents/MacOS
cp -r dist/TranslationAgent/* TranslationAgent.app/Contents/MacOS/

# 创建 Info.plist
cat > TranslationAgent.app/Contents/Info.plist << 'EOF'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>TranslationAgent</string>
    <key>CFBundleName</key>
    <string>Translation Agent</string>
    <key>CFBundleVersion</key>
    <string>1.0.0</string>
    <key>LSMinimumSystemVersion</key>
    <string>12.0</string>
</dict>
</plist>
EOF
```

### 8.3 Linux

```bash
# 直接运行
chmod +x dist/TranslationAgent
./dist/TranslationAgent

# 或者创建 .desktop 文件
cat > ~/.local/share/applications/translation-agent.desktop << 'EOF'
[Desktop Entry]
Name=Translation Agent
Comment=AI Translation Tool
Exec=/path/to/TranslationAgent
Icon=/path/to/app/image.png
Terminal=false
Type=Application
Categories=Utility;Office;
EOF
```

### 8.4 文件大小优化

如需减小体积：

1. **使用 UPX 压缩**（已默认在 onefile 中启用）：
   ```bash
   # 确保 UPX 可用
   upx --version
   ```

2. **排除不需要的模块**：在 `.spec` 的 `excludes` 中添加未使用的库。

3. **使用虚拟环境最小安装**：构建前只安装 `requirements-desktop.txt`，避免带入无关包。

---

## 附录：文件清单

| 文件 | 说明 |
|------|------|
| `requirements-desktop.txt` | 桌面版 Python 依赖清单 |
| `TranslationAgent-onefile.spec` | PyInstaller 单文件模式配置 |
| `TranslationAgent-onedir.spec` | PyInstaller 目录模式配置 |
| `build.bat` | Windows 一键构建脚本 |
| `build.sh` | macOS/Linux 一键构建脚本 |
| `.env.example` | 环境变量配置模板 |
| `packaging-guide.md` | 本文档 |

---

*最后更新：2025-01*
