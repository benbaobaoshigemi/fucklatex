@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

:: Word LaTeX公式自动转换器 - 启动脚本
:: 功能：先转换 $...$ 为公式对象，再自动打开Word并转换为专业格式

title Word LaTeX公式自动转换器

echo ========================================
echo  Word LaTeX公式自动转换器
echo  (一键完成：转换+打开+专业格式)
echo ========================================
echo.

:: 检查Python是否可用（直接使用conda环境）
set "PYTHON_CMD=C:/Users/zhang/miniconda3/python.exe"
echo ⏳ 正在测试Python...

:: 先测试直接使用conda目录下的python
%PYTHON_CMD% --version >nul 2>&1
if errorlevel 1 (
    echo ⚠️  直接Python不可用，使用conda run...
    set "PYTHON_CMD=C:/Users/zhang/miniconda3/Scripts/conda.exe run -p C:\Users\zhang\miniconda3 python"
    %PYTHON_CMD% --version >nul 2>&1
    if errorlevel 1 (
        echo ❌ 错误：Python不可用
        echo 💡 请检查conda安装是否正确
        pause
        exit /b 1
    )
)

echo ✅ Python检测通过

:: 获取拖放的文件或手动输入
set "INPUT_FILE=%~1"

if "%INPUT_FILE%"=="" (
    echo 💡 提示：可以直接拖拽Word文档到此脚本
    echo.
    set /p "INPUT_FILE=📂 请输入Word文档路径: "
)

:: 去除引号
set "INPUT_FILE=%INPUT_FILE:"=%"

:: 检查文件是否存在
if not exist "%INPUT_FILE%" (
    echo ❌ 错误：文件不存在
    echo 路径：%INPUT_FILE%
    pause
    exit /b 1
)

:: 检查是否为docx文件
if /i not "%INPUT_FILE:~-5%"==".docx" (
    echo ❌ 错误：只支持 .docx 格式的Word文档
    echo 当前文件：%INPUT_FILE%
    pause
    exit /b 1
)

:: 生成输出文件路径
for %%F in ("%INPUT_FILE%") do (
    set "FILE_DIR=%%~dpF"
    set "FILE_NAME=%%~nF"
    set "FILE_EXT=%%~xF"
)
set "OUTPUT_FILE=%FILE_DIR%%FILE_NAME%_converted%FILE_EXT%"

echo.
echo ========================================
echo 第一步：转换 $...$ 为公式对象
echo ========================================
echo.
echo 📄 输入文件: %INPUT_FILE%
echo 📄 输出文件: %OUTPUT_FILE%
echo.

:: 运行Python脚本转换LaTeX
set "PYTHONIOENCODING=utf-8"
%PYTHON_CMD% "%~dp0main.py" "%INPUT_FILE%" -o "%OUTPUT_FILE%" --auto-install

if errorlevel 1 (
    echo.
    echo ❌ 第一步失败！
    pause
    exit /b 1
)

echo.
echo ========================================
echo 第二步：打开Word并转换为专业格式
echo ========================================
echo.

:: 运行自动转换脚本
set "PYTHONIOENCODING=utf-8"
%PYTHON_CMD% "%~dp0auto_convert.py" "%OUTPUT_FILE%"

if errorlevel 1 (
    echo.
    echo ❌ 第二步失败！
    echo 💡 已生成转换文件，您可以手动在Word中操作
    echo 📂 文件位置: %OUTPUT_FILE%
    pause
    exit /b 1
)

echo.
echo ========================================
echo ✅ 全部完成！
echo ========================================
echo.
echo 📂 输出文件: %OUTPUT_FILE%
echo.
pause
