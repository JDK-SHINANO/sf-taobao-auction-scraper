#!/usr/bin/env python3
"""
拍卖信息 + 标的物介绍 合并工具
双击 run_merge.bat 运行，或把数据文件夹拖到 bat 文件上
"""

import os
import sys

TARGET_FILES = ["拍卖信息.txt", "基本信息.csv", "标的物介绍.csv"]


def find_items(root_dir):
    items = []
    for dirpath, dirnames, filenames in os.walk(root_dir):
        if any(f in filenames for f in TARGET_FILES):
            items.append((dirpath, os.path.basename(dirpath)))
    return items


def merge_items(items, output_file):
    count = 0
    with open(output_file, "w", encoding="utf-8") as out:
        for item_path, item_name in items:
            out.write("=" * 30 + "\n")
            out.write(f"项目： {item_name}\n")
            out.write("=" * 60 + "\n\n")
            for target in TARGET_FILES:
                filepath = os.path.join(item_path, target)
                if os.path.isfile(filepath):
                    with open(filepath, "r", encoding="utf-8") as f:
                        out.write(f.read())
                    out.write("\n")
            out.write("\n")
            count += 1
    return count


def main():
    print("=" * 50)
    print("  拍卖信息 + 标的物介绍 合并工具")
    print("=" * 50)
    print()

    # 确定目录：命令行参数 > 弹窗选择 > 脚本所在目录
    root_dir = None

    if len(sys.argv) > 1:
        # 拖拽文件夹到 bat 上时，文件夹路径作为参数传入
        root_dir = sys.argv[1]
    else:
        # 无参数：尝试弹窗选目录
        try:
            import tkinter as tk
            from tkinter import filedialog
            tk_root = tk.Tk()
            tk_root.withdraw()
            root_dir = filedialog.askdirectory(title="选择包含拍卖数据的文件夹")
            tk_root.destroy()
        except Exception:
            pass

        if not root_dir:
            # 弹窗失败或取消：用脚本所在目录
            root_dir = os.path.dirname(os.path.abspath(__file__))
            print(f"使用脚本所在目录: {root_dir}")

    if not os.path.isdir(root_dir):
        print(f"❌ 目录不存在: {root_dir}")
        input("\n按回车键退出...")
        return

    print(f"📂 扫描: {root_dir}")

    items = find_items(root_dir)
    if not items:
        print("❌ 未找到包含 拍卖信息.txt 或 标的物介绍.txt 的目录")
        input("\n按回车键退出...")
        return

    output_file = os.path.join(root_dir, "merged_output.txt")
    count = merge_items(items, output_file)
    size = os.path.getsize(output_file)

    print(f"✅ 完成！共处理 {count} 个标的物")
    print(f"📄 输出: {output_file} ({size / 1024:.1f} KB)")
    input("\n按回车键退出...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ 出错了: {e}")
        input("\n按回车键退出...")
