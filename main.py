import os
from pathlib import Path
import chardet
import json

print("=" * 30)
print(("|" + " " * 28 + "|\n") + ("|" +" " * 4 + "文件转换工具  v1.0.0" + " " * 4 + "|\n") + ("|" + " " * 28 + "|"))
print("=" * 30)

folder = Path(__file__).parent.resolve()

def ask_file_path():
    files = ["", ""]

    while True:
        files[0] = input("请输入想要转换的文件：")
        if os.path.exists(files[0]):
            break
        else:
            print("文件不存在，请重新输入！")

    while True:
        files[1] = input("请输入转换后的文件：")
        if os.path.exists(files[1]):
            print("文件已存在，请重新输入！")
        else:
            break

    return files[0], files[1]

def get_encoding(file_path):
    return chardet.detect(open(files[0], 'rb').read())['encoding']

while True:
    select_type = input("请选择文件类型\n1.压缩字符到1行 \n2.压缩json到1行 \n3.获取文件夹内所有快捷方式连接的路径\n4. 退出\n")
    if select_type in ["1", "2", "3", "4"]:
        break
    else:
        print("输入错误，请重新输入！")

if select_type in ["1", "2"]:
    files = ask_file_path()

if select_type == "1":
    with open(files[0], 'r', encoding=get_encoding(files[0])) as f:
        content = ' '.join(line.strip() for line in f)

    with open(files[1], 'w', encoding="utf-8") as f:
        f.write(content)
elif select_type == "2":
    with open(files[0], 'r', encoding=get_encoding(files[0])) as f:
        data = json.load(f)

    with open(files[1], 'w', encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
elif select_type == "3":
    os.system("powershell -ExecutionPolicy Bypass -File " + str(Path(folder, "copylink.ps1")))        
elif select_type == "4":
    exit()
    
os.system("pause")