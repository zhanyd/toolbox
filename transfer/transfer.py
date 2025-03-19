import os
import shutil
import time
import win32com.client
from win32com.client import constants

def convert_doc_to_docx(source_folder, target_folder):
    """
    将指定文件夹中的所有.doc文件转换为.docx文件，并保持文件夹结构不变。
    转换后的文件将保存到新建的目标文件夹中。
    使用Microsoft Word应用程序进行转换。
    """
    # 确保目标文件夹存在
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
    
    # 创建Word应用程序实例
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    try:
        # 遍历源文件夹
        for root, dirs, files in os.walk(source_folder):
            for file in files:
                # 跳过临时文件
                if file.startswith("~$"):
                    continue
                    
                if file.endswith('.doc'):
                    # 获取文件的完整路径
                    file_path = os.path.join(root, file)
                    # 创建目标文件夹的对应结构
                    relative_path = os.path.relpath(root, source_folder)
                    target_subfolder = os.path.join(target_folder, relative_path)
                    if not os.path.exists(target_subfolder):
                        os.makedirs(target_subfolder)
                    
                    # 构造目标文件路径（.docx扩展名）
                    target_file_path = os.path.join(target_subfolder, os.path.splitext(file)[0] + '.docx')
                    
                    try:
                        # 使用Word应用程序打开并另存为.docx
                        doc = word.Documents.Open(os.path.abspath(file_path))
                        # 使用数值16代替constants.wdFormatXMLDocument
                        doc.SaveAs(os.path.abspath(target_file_path), FileFormat=16)
                        doc.Close()
                        print(f"已转换 {file_path} 到 {target_file_path}")
                    except Exception as e:
                        print(f"转换 {file_path} 时出错: {e}")
    finally:
        # 确保Word应用程序被关闭
        word.Quit()

def convert_xls_to_xlsx(source_folder, target_folder):
    """
    将指定文件夹中的所有.xls文件转换为.xlsx文件，并保持文件夹结构不变。
    转换后的文件将保存到新建的目标文件夹中。
    使用Microsoft Excel应用程序进行转换。
    """
    # 确保目标文件夹存在
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
    
    # 创建Excel应用程序实例
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        # 遍历源文件夹
        for root, dirs, files in os.walk(source_folder):
            for file in files:
                # 跳过临时文件
                if file.startswith("~$"):
                    continue
                    
                if file.endswith('.xls'):
                    # 获取文件的完整路径
                    file_path = os.path.join(root, file)
                    # 创建目标文件夹的对应结构
                    relative_path = os.path.relpath(root, source_folder)
                    target_subfolder = os.path.join(target_folder, relative_path)
                    if not os.path.exists(target_subfolder):
                        os.makedirs(target_subfolder)
                    
                    # 构造目标文件路径（.xlsx扩展名）
                    target_file_path = os.path.join(target_subfolder, os.path.splitext(file)[0] + '.xlsx')
                    
                    try:
                        # 使用Excel应用程序打开并另存为.xlsx
                        workbook = excel.Workbooks.Open(os.path.abspath(file_path))
                        # 使用数值51代替constants.xlWorkbookDefault
                        workbook.SaveAs(os.path.abspath(target_file_path), FileFormat=51)
                        workbook.Close()
                        print(f"已转换 {file_path} 到 {target_file_path}")
                    except Exception as e:
                        print(f"转换 {file_path} 时出错: {e}")
    finally:
        # 确保Excel应用程序被关闭
        excel.Quit()

if __name__ == "__main__":
    # 源文件夹路径
    source_folder = "c:/文档/知识库/产品资料"
    # 目标文件夹路径（分别为Word和Excel文件）
    target_folder_doc = source_folder + "_docx"
    target_folder_xls = source_folder + "_xlsx"
    
    print("开始转换Word文档...")
    convert_doc_to_docx(source_folder, target_folder_doc)
    print("Word文档转换完成。")
    
    print("开始转换Excel文档...")
    convert_xls_to_xlsx(source_folder, target_folder_xls)
    print("Excel文档转换完成。")
    
    print("所有文件转换完成。")