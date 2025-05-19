import os
os.environ['TK_SILENCE_DEPRECATION'] = '1'
import re
import pdfplumber
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
from datetime import datetime

def extract_invoice_data(pdf_path):
    data = {
        "合同编号": "",
        "供应商名称": "",
        "项目名称": "",
        "发票金额": "",
        "开票日期": "",
        "发票号码": "",
        "数量": ""  # 新增数量字段
    }

    all_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text += text + "\n"

    lines = all_text.splitlines()
    quantities = []

    # 初始化 project_found 和 project_lines 以避免未赋值引用错误
    project_found = False
    project_lines = []

    for line in lines:
        line = line.strip()

        # 合同编号
        if not data["合同编号"] and 'SLG' in line:
            match = re.search(r"(SLG[\w\-]+)", line)
            if match:
                data["合同编号"] = match.group(1)

        # 供应商名称
        if not data["供应商名称"]:
            seller_match = re.search(r"销[\s\S]*?名称：\s*([^\n]+)", line)
            if seller_match:
                data["供应商名称"] = seller_match.group(1).strip()

        # 项目名称 提取：只提取第一个有效行，若第二行含*则不拼接
        if ("项目名称" in line or "货物名称" in line) and not project_found:
            project_found = True
            continue
        if project_found:
            original_line = line.strip()
            processed_line = re.sub(r"\*.*?\*", "", original_line).strip()

            if not project_lines:  # 处理第一行
                if processed_line and not any(bad in processed_line for bad in ["购方", "销方", "税号", "地址", "电话", "开户"]):
                    project_lines.append(processed_line.split()[0])
            else:  # 处理第二行
                # 如果原始行包含星号则终止拼接
                if '*' in original_line:
                    project_found = False
                else:
                    if processed_line and not any(bad in processed_line for bad in ["购方", "销方", "税号", "地址", "电话", "开户"]):
                        project_lines.append(processed_line.split()[0])
                project_found = False  # 无论是否拼接，最多取两行
            
            # 合并结果
            combined = "".join(project_lines)
            data["项目名称"] = re.sub(r"合$", "", combined)  # 新增替换逻辑

        # 开票日期
        if not data["开票日期"]:
            date_match = re.search(r"开票日期[:：]\s*(\d{4}年\d{2}月\d{2}日)", line)
            if date_match:
                data["开票日期"] = date_match.group(1)

        # 发票金额（优先匹配价税合计小写）
        if not data["发票金额"]:
            amount_match = re.search(r"价税合计.*?小写[^\n]*￥([\d,]+\.\d{2})", line)
            if not amount_match:
                amount_match = re.search(r"小写\s*￥([\d,]+\.\d{2})", line)
            if amount_match:
                data["发票金额"] = amount_match.group(1)

        # 发票号码
        if not data["发票号码"]:
            invoice_match = re.search(r"发票号码[:：]\s*(\d+)", line)
            if invoice_match:
                data["发票号码"] = invoice_match.group(1)

        # 提取数量（从表格行中匹配）
        if "*" in line:
            # 匹配格式：台    0.2  557522.123893805
            qty_match = re.search(r"(台|条|套|个)\s+([\d\.]+)\s", line)
            print(qty_match)
            if qty_match:
                quantity_str = qty_match.group(2)  # 提取数值部分
                try:
                    quantities.append(float(quantity_str))
                except ValueError:
                    print(f"警告：无法转换数值 '{quantity_str}'，跳过该行")

    # 取所有数量的最小值（或根据需要调整逻辑）
    if quantities:        
        data["数量"] = f"{int(min(quantities) * 100)}%"

    return data

def rename_pdfs(pdf_paths):
    for pdf_path in pdf_paths:
        data = extract_invoice_data(pdf_path)
        
        # 生成新文件名（包含数量）
        #new_name = f"发票_{data['合同编号']}_{data['发票号码']}_{data['供应商名称']}_{data['开票日期']}_{data['数量']}.pdf"
        new_name = f"发票{data['合同编号']}_{data['发票号码']}_{data['供应商名称']}_{data['项目名称']}_{data['数量']}.pdf"
        new_path = os.path.join(os.path.dirname(pdf_path), new_name)
        
        try:
            os.rename(pdf_path, new_path)
            print(f"重命名成功: {new_name}")
        except Exception as e:
            print(f"重命名失败: {e}")

def drop_pdf(event):
    import re
    # 支持 Windows 和 macOS 拖拽路径解析
    file_paths = re.findall(r'\{(.*?)\}', event.data)  # 提取大括号内的路径（Windows）
    if not file_paths:
        file_paths = event.data.strip().split()  # 回退：macOS 或无括号格式
    try:
        rename_pdfs(file_paths)
        messagebox.showinfo("成功", "文件重命名完成！")
    except Exception as e:
        messagebox.showerror("错误", f"重命名过程中出错：\n{str(e)}")

def select_pdfs():
    file_paths = filedialog.askopenfilenames(
        title="选择PDF文件",
        filetypes=[("PDF 文件", "*.pdf")],
    )
    if file_paths:
        try:
            rename_pdfs(file_paths)
            messagebox.showinfo("成功", "文件重命名完成！")
        except Exception as e:
            messagebox.showerror("错误", f"重命名过程中出错：\n{str(e)}")

def create_gui():
    root = TkinterDnD.Tk()
    root.title("PDF发票重命名工具")
    root.geometry("400x300")
    root.resizable(False, False)

    label = tk.Label(root, text="拖拽多个PDF文件到这里 ➔ 自动重命名", font=("微软雅黑", 12))
    label.pack(pady=40)

    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', drop_pdf)

    button = tk.Button(root, text="选择多个PDF文件进行重命名", font=("微软雅黑", 12), command=select_pdfs)
    button.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    create_gui()