import glob
import pathlib
import pikepdf
import pdfplumber
import re

"""
PDF 轉 TXT 檔

功能：
- 轉換 *.pdf 檔案格式
- 移除 PDF 檔案權限 (複製、編輯)
"""

def remove_permissions(filename):
    """ 移除 PDF 檔案權限 (複製、編輯) """
    file = f'{filename}.pdf'
    with pikepdf.open(file, allow_overwriting_input=True) as pdf:       # 使用 pikepdf 模組開啟 PDF       
        print(f"{file} remove permissions... ", end='')
        pdf.save(file)      # 複寫 PDF 檔 (移除權限)
        print("【Done】")

def save_txt(filename):
    """ 寫入 TXT 檔 """
    file = f'{filename}_py.txt'
    with open(file, 'a', encoding='utf-8') as f:        # 寫入 .txt 檔案 (a 附加模式)
        print(f"{file} create file... ", end='')
        with pdfplumber.open(f'{filename}.pdf') as pdf:     # 使用 pdfplumber 模組開啟 PDF
            # 開始寫入的關鍵字
            start_keywords = {
                '科目:共同科目(國文、英文)', 
                '科目:1.計算機原理 2.網路概論', 
                '科目:1.資訊管理 2.程式設計'}
            # 字首加換行的關鍵字
            line_keywords = {
                '壹、', '貳、', '參、',
                '一、', '二、', '三、'
                }

            write_flag = False      # 開始寫入的關鍵字 flag
            for page_number, page in enumerate(pdf.pages, start=1):     # pdfplumber 模組切換 PDF 頁面
                page = pdf.pages[page_number-1]
                text = page.extract_text()      # 擷取頁面文字
                lines = text.split('\n')

                for line in lines[:-1]:     # 去除 PDF 分頁提示，共同科目(國文、英文) 第 1 頁，共 4 頁 【 請 翻 頁 繼續作答】
                    if write_flag:      # 是否開始寫入 TXT
                        # 檢查行首是為「數字」或「換行的關鍵字」
                        if re.match(r'^\d+', line) or any(line.startswith(k) for k in line_keywords):
                            f.write('\n')       # 行首插入換行
                            #print('\n')
                        f.write(line + '\n')
                        #print(line)
                    elif any(k in line for k in start_keywords):        # 檢查該行是否為「寫入的關鍵字」
                        write_flag = True
                        f.write(f"\n{line}\n")      # 開始寫入首行
                        #print(line)
        print("【Done】")


if __name__ == '__main__':
    """ 主程式入口 """
    pdf_files = glob.glob('*.pdf')      # 取得當前目錄所有 *.pdf
    for file in pdf_files:
        filename = pathlib.Path(file).stem      # 取得檔案名稱，filename.pdf => filename

        txt_file = pathlib.Path(f'{filename}_py.txt')       # 檢查 _py.txt 檔案是否存在
        if txt_file.exists():
            txt_file.unlink()       # 刪除 .txt 檔案
        else:
            remove_permissions(filename)        # 移除 PDF 檔案權限 (複製、編輯)
        
        save_txt(filename)      # 寫入 TXT 檔
