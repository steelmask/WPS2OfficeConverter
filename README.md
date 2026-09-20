# WPS2OfficeConverter

图形界面工具：将 WPS 专有格式转换为 Office 通用格式。

| 源格式 | 输出格式 |
| --- | --- |
| WPS 表格 `.et` / `.ett` | Excel `.xlsx` / `.xls` |
| WPS 文字 `.wps` / `.wpt` | Word `.docx` / `.doc` |
| WPS 演示 `.dps` / `.dpt` | PowerPoint `.pptx` / `.ppt` |

## 使用前准备

推荐安装 **WPS Office（含 WPS 表格）**，它会通过 WPS 的自动化接口转换，兼容性最好。随后在本目录执行：

```powershell
py -m pip install -r requirements.txt
py et_converter.py
```

若电脑没有 WPS，工具会自动尝试使用 LibreOffice；请先安装 LibreOffice，并让 `soffice.exe` 可被系统找到。

## 使用方法

1. 点击“选择文件”，选取 WPS 专有格式文件。
2. 选择对应的 Office 输出格式。
3. 点击“转换并保存”，指定输出位置。

旧版 `.xls`、`.doc`、`.ppt` 格式存在功能限制；没有兼容旧系统的需求时，建议选择 `.xlsx`、`.docx`、`.pptx`。
