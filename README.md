# Excel 处理与比较工具

本项目是一个基于 Python Flask 和 pywebview 构建的桌面应用程序，旨在提供便捷的 Excel 文件处理和比较功能。

## 主要功能

1.  **Excel 文件批量处理**:
    *   **文件上传**: 支持拖放或点击选择多个 Excel 文件（`.xlsx`, `.xls`) 和 CSV 文件 (`.csv`)。
    *   **智能表头识别**: 自动检测并使用正确的表头行。
    *   **列名标准化**: 将不同的列名（中英文）映射到一套标准化的列名（如 `post_id`, `video_views`, `like` 等）。
    *   **数据清洗与规范化**:
        *   处理国内与国外数据文件，自动填充 `network` (平台) 字段。
        *   若 `profile` (账号) 信息缺失，尝试从文件名中提取。
        *   统一 `published_date` (发布日期) 格式为 `YYYY-MM-DD_HH:MM:SS`。
        *   根据 `post` (内容) 文本自动打上 `game_label` (游戏标签)。
        *   清理 `post` 文本中的多余空格和换行。
    *   **重复数据合并**: 若多个文件或同一文件内存在相同的 `post_id`，将合并这些记录，数值型字段（如 `video_views`, `like` 等）会进行累加。
    *   **结果下载**: 将所有处理后的数据合并到一个 Excel 文件中，供用户下载。

2.  **Excel 文件比较**:
    *   **双文件上传**: 用户上传一个旧版 Excel 文件和一个新版 Excel 文件。
    *   **基于 ID 比较**: 以 `post_id` 为基准，比较两个文件的数据差异。
    *   **差异计算**: 自动计算数值列（如 `video_views`）在新旧文件间的差值。
    *   **统计摘要**: 生成统计信息，包括帖子数量变化、各数值列的总计变化等。
    *   **比较结果下载**: 生成包含详细比较数据和统计摘要的 Excel 文件供用户下载。

3.  **用户友好的界面**:
    *   通过 `pywebview` 将 Flask Web 应用包装成桌面应用，提供原生体验。
    *   需要在浏览器打开 http://localhost:5001/ 才能使用

## 技术栈

*   **后端**: Python, Flask
*   **前端**: HTML, CSS, JavaScript
*   **数据处理**: Pandas, NumPy, openpyxl, xlrd
*   **GUI**: pywebview
*   **打包**: PyInstaller

## 如何运行 (开发模式)

1.  **环境准备**:
    *   确保已安装 Python 3 和 pip3。
    *   克隆或下载本仓库。

2.  **安装依赖**:
    在项目根目录下打开终端，运行：
    ```bash
    pip3 install -r requirements.txt
    ```

3.  **启动应用**:
    *   **macOS / Linux**:
        可以直接运行 `start.command` 脚本：
        ```bash
        ./start.command
        ```
        或者手动执行：
        ```bash
        python3 app.py
        ```
    *   **Windows**:
        ```bash
        python app.py
        ```
    应用将在一个桌面窗口中打开，默认访问 `http://localhost:5001`。

## 如何构建可执行文件

本项目使用 PyInstaller 进行打包。

### 1. macOS

*   **安装 PyInstaller** (如果尚未安装):
    ```bash
    pip3 install pyinstaller
    ```
*   **执行打包**:
    在项目根目录下运行：
    ```bash
    pyinstaller Excel处理工具.spec
    ```
    或者，如果 `Excel处理工具.spec` 文件配置了 `console=False` (针对 GUI 应用):
    ```bash
    pyinstaller --noconfirm --windowed Excel处理工具.spec
    ```
    打包完成后，可在 `dist` 目录下找到 `Excel处理工具.app`。

### 2. Windows (通过 Docker 交叉编译)

如果需要在非 Windows 环境下为 Windows 打包，可以使用提供的 `Dockerfile.windows`。

*   **安装 Docker**: 确保你的系统已安装 Docker。
*   **构建 Docker 镜像并打包**:
    在项目根目录下运行：
    ```bash
    docker build -t excel-tool-builder -f Dockerfile.windows .
    ```
*   **提取可执行文件**:
    打包成功后，可执行文件 `Excel处理工具.exe` 会在 Docker 镜像内部的 `/src/dist/` 目录。你需要从 Docker 容器中将其复制出来。
    例如，创建一个临时容器并复制文件：
    ```bash
    docker create --name temp_container excel-tool-builder
    docker cp temp_container:/src/dist/Excel处理工具.exe ./
    docker rm temp_container
    ```
    之后，`Excel处理工具.exe` 文件就会出现在你的项目根目录。

### 3. Windows (本地环境)

*   **安装 PyInstaller** (如果尚未安装):
    ```bash
    pip install pyinstaller
    ```
*   **执行打包**:
    在项目根目录下运行：
    ```bash
    pyinstaller Excel处理工具.spec
    ```
    或者，如果 `Excel处理工具.spec` 文件配置了 `console=False` (针对 GUI 应用):
    ```bash
    pyinstaller --noconfirm --windowed Excel处理工具.spec
    ```
    打包完成后，可在 `dist` 目录下找到 `Excel处理工具.exe` (如果 `onefile` 选项启用) 或一个包含可执行文件和依赖的文件夹。

## 文件结构说明

*   `app.py`: Flask 应用主程序，处理 HTTP 请求和 GUI 逻辑。
*   `main.py`: 包含核心的 Excel 文件处理逻辑 ( `process_excel` 函数等)。
*   `compare_excel.py`: 包含 Excel 文件比较逻辑 ( `compare_excel_files` 函数)。
*   `requirements.txt`: 项目依赖的 Python 包。
*   `start.command`: macOS/Linux 下快速启动脚本。
*   `Excel处理工具.spec`: PyInstaller 打包配置文件。
*   `Dockerfile.windows`: 用于 Docker 构建 Windows 可执行文件的配置。
*   `templates/`:存放 HTML 模板。
    *   `index.html`: 应用主页面。
*   `static/`: 存放 CSS、JavaScript 和图片等静态资源。
    *   `css/style.css`: 主要样式表。
    *   `js/main.js`: 前端交互逻辑。
*   `.github/workflows/build.yml`: (如果存在) GitHub Actions 配置文件，可能用于自动化构建。

## 注意事项

*   上传的文件会临时存储在系统的临时目录中，并在处理完成后或应用关闭后尝试清理。
*   处理大文件时可能需要一些时间，请耐心等待。