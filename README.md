# MultiDoc Sync

一个本地、只读、开源的多文档并排联动工具。一次选择 3～4 个同类文件，即可在同一屏幕上并排比较。

## 特性

- PDF：直接在内置阅读器中显示，不转换、不上传。
- Word：直接调用 Microsoft Word 打开原文件，不转换；Windows 支持滚动位置联动。
- PowerPoint：直接调用 Microsoft PowerPoint 打开原文件，不转换；Windows 支持相同幻灯片编号联动。
- 每一栏都可以单独解除联动。
- 三文件三列、四文件四列，并适配不同显示器和缩放比例。
- 原文件只读打开，不写入、不覆盖。

每次必须选择同一类文件：PDF；Word（DOC/DOCX 可混用）；或 PowerPoint（PPT/PPTX 可混用）。

## 下载与运行

在 [Releases](https://github.com/shenquanzhen/multidoc-sync/releases) 下载对应平台版本：

- Windows：`MultiDocSync-Windows-x64.exe`，无需安装，双击运行。
- macOS Apple Silicon：`MultiDocSync-macOS-arm64.zip`，解压后打开应用。

PDF 模式不需要外部软件。Word/PPT 原生模式需要本机安装 Microsoft Word 或 PowerPoint；缺失时程序会给出提示。程序不会捆绑 Office 或 LibreOffice。

> macOS 说明：PDF 功能完整；Word/PPT 可以原生打开和排列，并可用控制栏同步上一页/下一页。该功能依赖 macOS“辅助功能”权限。由于 Office for Mac 没有稳定公开的跨窗口滚动状态接口，暂不提供鼠标滚轮自动跟随。

## 从源码运行

```bash
python -m pip install -r requirements.txt
python multi_document_viewer.py
```

Windows 也可以双击 `启动多文档对比.cmd`。如果缺少 Python 依赖，脚本会显示安装提示。

## 构建

Windows PowerShell：

```powershell
./scripts/build_windows.ps1
```

macOS：

```bash
bash scripts/build_macos.sh
```

PyInstaller 不是跨平台编译器，因此 Windows 与 macOS 包必须在各自系统上构建。仓库中的 GitHub Actions 会分别完成构建并提供下载产物。

## 隐私与安全

- 所有文档只在本机处理。
- 不上传文件，不读取浏览器数据；匿名统计默认关闭，只有首次明确同意后才发送最少量数据。
- Office 文件以只读方式打开，并禁用宏自动化。
- 发布前由自动检查阻止常见密钥、Cookie、令牌和本机路径进入仓库。
- 匿名使用统计只有在首次明确同意后才启用；详见 [PRIVACY.md](PRIVACY.md)。
- 统计接收端可自托管在阿里云，示例位于 [server](server/README.md)。

仓库源码中的统计地址默认为空，因此直接从源码构建不会发送任何统计。部署自有接收端并填写公开的 HTTPS 接口地址后，发行版才会在首次运行时询问用户是否同意。

## 许可证

[MIT](LICENSE)
