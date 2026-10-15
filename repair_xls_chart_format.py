import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client as win32


def repair_excel_charts(folder_path: str, backup: bool = True):
    xls_files = [
        f for f in os.listdir(folder_path)
        if f.lower().endswith(".xls") and not f.startswith("~$")
    ]

    if not xls_files:
        messagebox.showinfo("提示", "当前文件夹下没有找到 .xls 文件。")
        return

    if backup:
        backup_dir = os.path.join(folder_path, "backup_original_xls")
        os.makedirs(backup_dir, exist_ok=True)

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    success_count = 0
    fail_list = []

    try:
        for file_name in xls_files:
            file_path = os.path.join(folder_path, file_name)

            try:
                if backup:
                    shutil.copy2(file_path, os.path.join(backup_dir, file_name))

                wb = excel.Workbooks.Open(file_path)

                for ws in wb.Worksheets:
                    for chart_obj in ws.ChartObjects():
                        chart = chart_obj.Chart

                        # 修复坐标轴格式
                        try:
                            chart.Axes(1).TickLabels.NumberFormat = "General"  # 横轴
                        except Exception:
                            pass

                        try:
                            chart.Axes(2).TickLabels.NumberFormat = "General"  # 纵轴
                        except Exception:
                            pass

                        # 修复趋势线公式和 R² 标签格式
                        try:
                            for series in chart.SeriesCollection():
                                for trendline in series.Trendlines():
                                    try:
                                        trendline.DataLabel.NumberFormat = "General"
                                    except Exception:
                                        pass
                        except Exception:
                            pass

                wb.Save()
                wb.Close(SaveChanges=False)

                success_count += 1

            except Exception as e:
                fail_list.append(f"{file_name}：{e}")

    finally:
        excel.Quit()

    msg = f"处理完成！\n\n成功处理：{success_count} 个文件"

    if fail_list:
        msg += f"\n失败：{len(fail_list)} 个文件\n\n"
        msg += "\n".join(fail_list[:10])

    messagebox.showinfo("完成", msg)


def main():
    root = tk.Tk()
    root.withdraw()

    folder_path = filedialog.askdirectory(title="请选择包含 xls 文件的文件夹")

    if not folder_path:
        return

    confirm = messagebox.askyesno(
        "确认",
        f"即将处理该文件夹下所有 .xls 文件：\n\n{folder_path}\n\n"
        "程序会先自动备份原文件，然后修复并覆盖原 xls。\n\n是否继续？"
    )

    if not confirm:
        return

    repair_excel_charts(folder_path, backup=True)


if __name__ == "__main__":
    main()