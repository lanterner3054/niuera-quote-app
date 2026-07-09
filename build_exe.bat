@echo off
REM ── 把报价工作台打包成单个 .exe (在 Windows 上运行) ──
REM 先装依赖与打包工具:
pip install -r requirements.txt pyinstaller
REM 打包(把 data/ 与 static/ 一并塞进 exe;launcher.py 启动服务并开浏览器):
pyinstaller --noconfirm --onefile --name NIUERA报价工作台 ^
  --add-data "static;static" ^
  --add-data "data;data" ^
  --add-data "engine;engine" ^
  --collect-all uvicorn --collect-all fastapi ^
  launcher.py
echo.
echo 打包完成,exe 在 dist\ 目录。双击即用。
echo 注意:data\ 里的 products.json / config.json 已被打进 exe;
echo 若想让客户端可随时改产品,改用 --add-data 外置数据(见 README)。
pause
