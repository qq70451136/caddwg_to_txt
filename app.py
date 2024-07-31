from pyautocad import Autocad, APoint
import os
import re
import time
from pyautocad.utils import unformat_mtext

# 正则表达式用于匹配 .dwg 文件
reg = re.compile(r'.*(.dwg)$')

# 正则表达式用于检测纯数字文本
pure_number_reg = re.compile(r'^\d+$')

# 指定路径
path = r"E:\python\dabao"
dwg_files = []

# 遍历指定目录，获取所有 .dwg 文件的路径
for dirpath, dirname, filename in os.walk(path):
    files = [os.path.join(dirpath, s) for s in filename if os.path.isfile(os.path.join(dirpath, s)) if reg.findall(s)]
    dwg_files.extend(files)

# 检查是否找到了任何dwg文件
if not dwg_files:
    print("No DWG files found in the specified directory.")
    exit()

# 连接到AutoCAD
acad = Autocad(create_if_not_exists=True, visible=True)

# 定义一个函数来确保关闭所有打开的文档
def close_all_documents():
    try:
        while acad.app.Documents.Count > 0:
            acad.ActiveDocument.Close(False)
    except Exception as e:
        print(f"Failed to close document: {e}")

# 初始化操作（只执行一次）
print("Initializing...")  # 打印初始化信号
placeholder_dwg = dwg_files[0]
try:
    acad.app.Documents.Open(placeholder_dwg)
    time.sleep(3)  # 等待3秒，可以根据实际情况调整
    close_all_documents()
except Exception as e:
    print(f"Failed to open or close placeholder DWG: {e}")
    exit()

# 处理所有找到的 .dwg 文件
for filename in dwg_files:
    try:
        # 确保在打开新的文档之前关闭所有打开的文档
        close_all_documents()

        # 打开DWG文件
        acad.app.Documents.Open(filename)
        print("Hello! Extracting text from: {}\n".format(filename))

        # 等待文档完全打开
        time.sleep(2)  # 等待2秒，可以根据实际情况调整

        # 遍历CAD文件中的所有单行文本对象，打印其文本内容
        for entity in acad.iter_objects("Text"):
            if entity.ObjectName == 'AcDbText':  # 确保是单行文本
                text_content = entity.TextString
                if not pure_number_reg.match(text_content):  # 排除纯数字文本
                    print("Text: {}".format(text_content))

        # 遍历CAD文件中的所有多行文本对象，打印其文本内容
        for entity in acad.iter_objects("MText"):
            if entity.ObjectName == 'AcDbMText':  # 确保是多行文本
                mtext_content = unformat_mtext(entity.TextString)  # 提取多行文本，去掉格式信息
                if not pure_number_reg.match(mtext_content):  # 排除纯数字文本
                    print("MText: {}".format(mtext_content))

    except Exception as e:
        print(f"Error processing file {filename}: {e}")

    finally:
        # 确保在处理完当前文档后关闭所有打开的文档
        close_all_documents()