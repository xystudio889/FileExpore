# 导入库模块
import os
import chardet
import json
from win32com.client import Dispatch
import glob

# 显示标题和更新日志
print("=" * 30)
print(("|" + " " * 28 + "|\n") + ("|" +" " * 4 + "文件转换工具  v1.1.0" + " " * 4 + "|\n") + ("|" + " " * 28 + "|"))
print("=" * 30)
print("v1.1.1 更新日志：回退了gui的文件选择框")
print("\n")
print("下个版本预告：添加批量处理模式")

# 定义函数
def get_encoding(file_path):
    return chardet.detect(open(file_path, 'rb').read())['encoding']

def ask_text(title:str, asks:dict[str, str]):
    text = title
    for k, v in asks.items():
        text += f"\n{k}.{v}"
    while True:
        select_type = input(text + "\n请选择:")
        if select_type in asks.keys():
            return select_type
        else:
            print("输入错误，请重新输入！")

def ask_file(*texts, is_save: tuple[bool] = (False,)):
    if len(texts) != len(is_save):
        raise ValueError("texts参数个数与is_save参数个数不匹配")
    
    files = []
    for prompt, save in zip(texts, is_save):
        while True:
            file = input(f"{prompt}: ").strip()
            if file:
                if not save:
                    if not os.path.exists(file):
                        print("输入的文件不存在，请重新输入！")
                    elif os.path.isdir(file):
                        print("输入的是文件夹，请重新输入！")
                    else:
                        files.append(file)
                        break
                else:
                    if os.path.isdir(file):
                        print("输入的是文件夹，请重新输入！")
                        continue
                    if os.path.exists(file):
                        save_type = ask_text("文件已存在，是否覆盖？", {"1": "覆盖", "2": "跳过"})
                        if save_type == "1":
                            files.append(file)
                            break
                        elif save_type == "2":
                            raise ValueError("被用户取消")
                    else:
                        files.append(file)
                        break
            else:
                print("未选择文件，请重新输入！")
    
    return files if len(files) > 1 else files[0]

def ask_dir(*texts, is_save: tuple[bool]=(False,)):
    if len(texts) != len(is_save):
        raise ValueError("texts参数个数与is_save参数个数不匹配")
    
    dirs = []
    for prompt, save in zip(texts, is_save):
        while True:
            dir_path = input(f"{prompt}: ").strip()
            if dir_path:
                if not save:
                    if not os.path.exists(dir_path):
                        print("输入的文件夹不存在，请重新输入！")
                    elif not os.path.isdir(dir_path):
                        print("输入的是文件，请重新输入！")
                    else:
                        dirs.append(dir_path)
                        break
                else:
                    if os.path.isfile(dir_path):
                        print("输入的是文件，请重新输入！")
                        continue
                    if os.path.exists(dir_path):
                        save_type = ask_text("文件夹已存在，是否覆盖？", {"1": "覆盖", "2": "跳过"})
                        if save_type == "1":
                            dirs.append(dir_path)
                            break
                        elif save_type == "2":
                            raise ValueError("被用户取消")
                    else:
                        dirs.append(dir_path)
                        break
            else:
                print("未选择文件夹，请重新输入！")
    
    return dirs if len(dirs) > 1 else dirs[0]


# 主程序
def main():
    while True:
        select_type = ask_text("文件转换工具", {"1": "将文件中的多行字符合并为1行", "2": "格式化json", "3": "获取文件夹内所有快捷方式连接的路径", "4": "通过文件读取的命令行参数启动", "5": "退出"})
        if select_type == "1":
            try:
                files = ask_file("请选择输入文件", "请选择输出文件", is_save=(False, True))
            except Exception as e:
                print(f'文件读取失败：{e}')
                continue
            with open(files[0], 'r', encoding=get_encoding(files[0])) as f:
                content = ' '.join(line.strip() for line in f)

            with open(files[1], 'w', encoding="utf-8") as f:
                f.write(content)
            print(f"文件转换完成！放置在{files[1]}")
        elif select_type == "2":
            format_type = ask_text("格式化json", {"1": "压缩为1行", "2": "普通格式化", "3": "退出"})
            try:
                files = ask_file("请选择输入文件", "请选择输出文件", is_save=(False, True))
            except Exception as e:
                print(f'文件读取失败：{e}')
                continue
            try:
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
            except Exception as e:
                print(f"文件转换失败：{e}")

        elif select_type == "3":
            link_type = ask_text("获取快捷方式连接的路径", {"1": "文件快捷方式", "2": "url快捷方式", "3": "退出"})
            try:
                input_folder = ask_dir("请输入快捷方式文件夹路径") 
                output_file = ask_file("请输入输出文件路径", is_save=(True,))
            except Exception as e:
                print(f'文件读取失败：{e}')
                continue
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
            try:
                program = ask_file("请输入程序路径")
                command_file = ask_file("请输入命令文件")
            except Exception as e:
                print(f'文件读取失败：{e}')
                continue
            
            with open(command_file, 'r', encoding=get_encoding(command_file)) as f:
                command = f.read()
            os.system(f'"{program}" {command}')
        elif select_type == "5":
            break
        
        continue_processing = ask_text("是否继续处理？", {"1": "继续", "2": "退出"})
        if continue_processing == "2":
            break

try:
    main()
except (KeyboardInterrupt, EOFError):
    print("\n程序退出")