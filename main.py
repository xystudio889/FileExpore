# 导入库模块
import os
from pathlib import Path
import chardet
import json
from tkinter import filedialog
from win32com.client import Dispatch
import glob

# 显示标题和更新日志
print("=" * 30)
print(("|" + " " * 28 + "|\n") + ("|" +" " * 4 + "文件转换工具  v1.1.0" + " " * 4 + "|\n") + ("|" + " " * 28 + "|"))
print("=" * 30)
print("v1.1.0更新日志：\n1.json文件处理添加格式化\n2.添加批量获取url快捷方式的网址功能\n3.添加了部分功能二级输入，修改了部分提示\n4.获取文件部分变为图形窗口选择\n5.添加通过文件读取的命令行参数启动\n6.'获取文件夹内所有快捷方式连接的路径'功能不再依赖powershell,转为完全python实现")
print("\n")
print("下个版本预告：添加批量处理模式")
print("\n")

folder = Path(__file__).parent.resolve()

# 定义函数
def get_encoding(file_path):
    return chardet.detect(open(file_path, 'rb').read())['encoding']

def ask_files(*file_type):
    input_files = filedialog.askopenfilename(title="请选择文件", filetypes=[*file_type])
    if not input_files:
        print("未选择文件，程序退出……")
        exit()
    output_files = filedialog.asksaveasfilename(title="请选择输出文件", filetypes=[*file_type])
    if not output_files:
        print("未选择文件，程序退出……")
        exit()
    return input_files, output_files

# 主程序
while True:
    while True:
        select_type = input("请选择文件类型\n1.压缩字符到1行 \n2.格式化json \n3.获取文件夹内所有快捷方式连接的路径\n4.通过文件读取的命令行参数启动\n5. 退出\n")
        if select_type in ["1", "2", "3", "4", "5"]:
            break
        else:
            print("输入错误，请重新输入！")

    if select_type == "1":
        files = ask_files(("文本文件", "*.txt"), ("所有文件", "*.*"))
        with open(files[0], 'r', encoding=get_encoding(files[0])) as f:
            content = ' '.join(line.strip() for line in f)

        with open(files[1], 'w', encoding="utf-8") as f:
            f.write(content)
        print(f"文件转换完成！放置在{files[1]}")
    elif select_type == "2":
        while True:
            format_type = input("请选择格式化类型\n1.压缩为1行 \n2.普通格式化 \n3.退出\n")
            if format_type in ["1", "2", "3"]:
                break
            else:
                print("输入错误，请重新输入！")
        files = ask_files(("JSON文件", "*.json"))
        if format_type == "1":
            with open(files[0], 'r', encoding=get_encoding(files[0])) as f:
                data = json.load(f)

            with open(files[1], 'w', encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
        elif format_type == "2":
            with open(files[0], 'r', encoding=get_encoding(files[0])) as f:
                data = json.load(f)
            
            with open(files[1], 'w', encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
        elif format_type == "3":
            continue
        print(f"文件转换完成！放置在{files[1]}")

    elif select_type == "3":
        while True:
            link_type = input("请选择快捷方式类型\n1.文件快捷方式 \n2.url快捷方式 \n3.退出\n")
            if link_type in ["1", "2", "3"]:
                break
            else:
                print("输入错误，请重新输入！")
        input_folder = filedialog.askdirectory(title="请选择文件夹")
        if not input_folder:
            print("未选择文件夹，程序退出……")
            break
        
        output_file = filedialog.asksaveasfilename(title="请选择输出文件", filetypes=[("文本文件", "*.txt")])
        if not output_file:
            print("未选择文件，程序退出……")
            break
        if link_type == "1":
            files = glob.glob(os.path.join(input_folder, "*.lnk"))
            if not files:
                print("该目录下未找到.lnk文件")
            else:
                targets = {}
                for root, _, files in os.walk(input_folder):
                    for file in files:
                        if file.endswith('.lnk'):
                            lnk_path = os.path.join(root, file)
                            try:
                                shell = Dispatch('WScript.Shell')
                                shortcut = shell.CreateShortCut(lnk_path)
                                targets[lnk_path] = shortcut.Targetpath
                            except Exception as e:
                                print(f"解析失败 {lnk_path}: {str(e)}")
                with open(output_file, 'w', encoding="utf-8") as f:
                    f.write("\n".join(list(targets.values())))
        elif link_type == "2":
            files = glob.glob(os.path.join(input_folder, "*.url"))
            url = []

            if not files:
                print("该目录下未找到.url文件")
            else:
                for file in files:
                    try:
                        with open(file, 'r', encoding='utf-8-sig') as f:  # 处理BOM字符
                            for line in f:
                                if line.strip().startswith('URL='):
                                    url.append(line.split('=', 1)[1].strip())
                                    continue
                    except Exception as e:
                        print(f"错误：无法读取文件 {file} - {str(e)}")
            
            with open(output_file, 'w', encoding="utf-8") as f:
                f.write("\n".join(url))
        elif link_type == "3":
            continue
        print(f"文件转换完成！放置在{output_file}")
    elif select_type == "4":
        program = filedialog.askopenfilename(title="请选择程序", filetypes=[("可执行文件", "*.exe"), ("批处理文件", "*.bat *.cmd"), ("所有文件", "*.*")])
        if not program:
            print("未选择文件，程序退出……")
            break
        command = filedialog.askopenfilename(title="请选择程序启动参数文件", filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")])
        if not command:
            print("未选择文件，程序退出……")
            break
        os.system(f"{program} {command}")
    elif select_type == "5":
        break
    while True:
        continue_processing = input("是否继续处理？\n1.继续 \n2.退出\n")
        if continue_processing in ["1", "2"]:
            break
        else:
            print("输入错误，请重新输入！")
    if continue_processing == "2":
        break