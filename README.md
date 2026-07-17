# BOL Windows App

这是把你当前 `n8n` 流程改造成 Windows 桌面应用的 Python 源码版本。

## 功能

- 登录信息写死在代码里
- 界面只保留 `page` 输入
- 输入 `page=2` 时，会抓取第 `1` 页和第 `2` 页并合并输出
- 带进度条和运行日志
- 导出 Excel 后，可直接点击按钮打开结果文件

## 先改配置

打开 [config.py](./config.py)，填写：

- `CLIENT_ID`
- `CLIENT_SECRET`

## 本地运行

```bash
pip install -r requirements.txt
python app.py
```

## 打包成 Windows EXE

建议使用 `PyInstaller`：

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --name BOLShipmentApp app.py
```

打包后可执行文件会出现在 `dist/BOLShipmentApp.exe`。
