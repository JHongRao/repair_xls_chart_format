# Excel 图表格式批量修复工具

通过 Windows 上安装的 Microsoft Excel，批量将 `.xls` 文件内嵌图表的横纵坐标轴刻度标签、趋势线公式及 R² 标签的数字格式设为“常规”（`General`）。

## 使用

需要 Windows、Microsoft Excel 和 Python（含 tkinter）。安装依赖并运行：

```powershell
python -m pip install -r requirements.txt
python repair_xls_chart_format.py
```

选择包含 `.xls` 文件的文件夹，确认后开始处理。程序只扫描所选目录，不递归扫描子目录，并跳过 `~$` 开头的临时文件。

原文件会先复制到所选目录中的 `backup_original_xls` 文件夹，然后在原路径保存修改。重复处理会覆盖同名备份；需要长期保留的原件请另行保存。

## 范围与限制

- 只处理 `.xls` 文件，不处理 `.xlsx`、`.xlsm` 等格式。
- 只遍历工作表内嵌图表，不处理独立图表工作表或次坐标轴。
- 部分标签或坐标轴修改异常会被跳过；成功计数表示文件已完成保存，不保证每个图表属性都已修改。
- 处理完成后显示成功文件数及最多十条失败信息。

## 打包

```powershell
python -m pip install pyinstaller
python -m PyInstaller "Excel图表格式批量修复工具.spec"
```

生成的程序位于 `dist` 目录，运行时仍需要本机安装 Microsoft Excel。

仓库不包含本地数据、备份、构建产物或可执行文件。
