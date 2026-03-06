#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
隐患整改台账与报告生成工具
========================

用于处理安全生产管理中的隐患整改通知单、台账和报告生成的工具。
支持从 Excel 读取数据，结合图片压缩包，生成带图片的 Excel 台账和 Word 报告。

主要功能:
    - 将隐患照片和闭环照片嵌入 Excel 台账
    - 根据模板生成检查报告
    - 根据模板生成闭环报告

作者：Safety Management Team
版本：2.0.0
"""

import os
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional, Tuple, Any
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
import io
import warnings

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from datetime import datetime
import pandas as pd

# ============================================================================
# HEIC 支持配置
# ============================================================================
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
except ImportError:
    pass

warnings.filterwarnings("ignore", category=UserWarning, module='PIL')

# ============================================================================
# 常量定义
# ============================================================================

# 图片尺寸配置 (单位：厘米)
HAZARD_PHOTO_WIDTH_CM = 5.0
HAZARD_PHOTO_HEIGHT_CM = 3.5

# DPI 设置
IMAGE_DPI = 600
CM_TO_PIXEL = IMAGE_DPI / 2.54

# 计算像素尺寸
IMG_WIDTH_PX = int(HAZARD_PHOTO_WIDTH_CM * CM_TO_PIXEL)
IMG_HEIGHT_PX = int(HAZARD_PHOTO_HEIGHT_CM * CM_TO_PIXEL)

# Excel 行列配置
COL_WIDTH_FOR_IMG = 35
ROW_HEIGHT_FOR_IMG = 100

# Word 图片尺寸 (英寸)
WORD_IMG_WIDTH_INCHES = 1.97  # 5cm
WORD_IMG_HEIGHT_INCHES = 1.38  # 3.5cm

# 字体配置
FONT_NAME = '仿宋'
FONT_SIZE_PT = 14  # 四号字

# 必需的 Excel 列
REQUIRED_EXCEL_COLUMNS = [
    '隐患编号', '异常类别', '隐患级别', '异常事项',
    '班组', '整改人', '发现时间', '要求闭环时间'
]

# ZIP 文件夹名称
FOLDER_HAZARD_PHOTOS = "隐患照片"
FOLDER_CLOSE_LOOP_PHOTOS = "闭环照片"


# ============================================================================
# 图片处理模块
# ============================================================================

def resize_image_to_buffer(
    img_data: bytes,
    width_px: int = IMG_WIDTH_PX,
    height_px: int = IMG_HEIGHT_PX,
    dpi: int = IMAGE_DPI,
    quality: int = 95
) -> io.BytesIO:
    """
    调整图片尺寸并保存到缓冲区
    
    Args:
        img_data: 原始图片数据
        width_px: 目标宽度 (像素)
        height_px: 目标高度 (像素)
        dpi: 输出 DPI
        quality: JPEG 质量 (1-100)
    
    Returns:
        包含调整后图片的 BytesIO 对象
    """
    pil_img = PILImage.open(io.BytesIO(img_data))
    pil_img = pil_img.resize((width_px, height_px), PILImage.LANCZOS)
    
    if pil_img.mode in ("RGBA", "P"):
        pil_img = pil_img.convert("RGB")
    
    img_buffer = io.BytesIO()
    pil_img.save(img_buffer, format='JPEG', dpi=(dpi, dpi), quality=quality)
    img_buffer.seek(0)
    
    return img_buffer


def extract_zip_to_dict(zip_path: str) -> Dict[str, Dict[str, bytes]]:
    """
    从 ZIP 文件提取图片到字典
    
    Args:
        zip_path: ZIP 文件路径
    
    Returns:
        包含两个键的字典：'隐患照片' 和 '闭环照片'，每个键对应一个 {文件名：图片数据} 的字典
    
    Raises:
        FileNotFoundError: ZIP 文件不存在
        zipfile.BadZipFile: ZIP 文件损坏
    """
    folder_images = {FOLDER_HAZARD_PHOTOS: {}, FOLDER_CLOSE_LOOP_PHOTOS: {}}
    required_folders = {FOLDER_HAZARD_PHOTOS, FOLDER_CLOSE_LOOP_PHOTOS}
    
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        all_files = zip_ref.namelist()
        top_dirs = set(f.split('/')[0] for f in all_files if '/' in f)
        
        for folder in required_folders:
            if folder not in top_dirs:
                print(f"警告：ZIP 中缺少 '{folder}' 文件夹，将跳过相关处理。")
                continue
            
            for f in all_files:
                if f.startswith(folder + '/') and not f.endswith('/'):
                    filename = os.path.basename(f)
                    key = os.path.splitext(filename)[0]
                    with zip_ref.open(f) as img_file:
                        folder_images[folder][key] = img_file.read()
    
    return folder_images


# ============================================================================
# Excel 处理模块
# ============================================================================

def validate_excel_header(ws) -> None:
    """
    验证 Excel 表头是否正确
    
    Args:
        ws: openpyxl worksheet 对象
    
    Raises:
        ValueError: 表头不正确
    """
    if ws.cell(row=1, column=1).value != "隐患编号":
        raise ValueError('Excel 第一列标题必须是"隐患编号"')


def insert_image_to_cell(
    ws,
    cell_address: str,
    img_data: bytes,
    row_idx: int
) -> Tuple[bool, Optional[str]]:
    """
    向 Excel 单元格插入图片
    
    Args:
        ws: openpyxl worksheet 对象
        cell_address: 单元格地址 (如 "E5")
        img_data: 图片数据
        row_idx: 行索引
    
    Returns:
        (成功标志，错误信息)
    """
    try:
        img_buffer = resize_image_to_buffer(img_data)
        xl_img = XLImage(img_buffer)
        xl_img.width = IMG_WIDTH_PX
        xl_img.height = IMG_HEIGHT_PX
        ws.add_image(xl_img, cell_address)
        return True, None
    except Exception as e:
        return False, str(e)


def embed_images_to_excel(
    excel_path: str,
    zip_path: str,
    output_path: str
) -> List[str]:
    """
    将图片嵌入 Excel 台账
    
    Args:
        excel_path: 输入 Excel 文件路径
        zip_path: 图片 ZIP 文件路径
        output_path: 输出 Excel 文件路径
    
    Returns:
        错误信息列表
    
    Raises:
        ValueError: Excel 格式不正确
        FileNotFoundError: 文件不存在
    """
    wb = load_workbook(excel_path)
    ws = wb.active
    
    validate_excel_header(ws)
    folder_images = extract_zip_to_dict(zip_path)
    
    errors = []
    
    # 设置列宽
    ws.column_dimensions['E'].width = COL_WIDTH_FOR_IMG
    ws.column_dimensions['M'].width = COL_WIDTH_FOR_IMG
    
    row_idx = 2
    while True:
        cell_value = ws.cell(row=row_idx, column=1).value
        if cell_value is None:
            break
        
        key = str(cell_value).strip()
        inserted = False
        
        # 插入隐患照片 (E 列)
        if key in folder_images[FOLDER_HAZARD_PHOTOS]:
            success, error = insert_image_to_cell(
                ws, f"E{row_idx}",
                folder_images[FOLDER_HAZARD_PHOTOS][key],
                row_idx
            )
            if success:
                inserted = True
            else:
                errors.append(f"第 {row_idx} 行 (隐患编号 {key}) 隐患照片插入失败：{error}")
        
        # 插入闭环照片 (M 列)
        if key in folder_images[FOLDER_CLOSE_LOOP_PHOTOS]:
            success, error = insert_image_to_cell(
                ws, f"M{row_idx}",
                folder_images[FOLDER_CLOSE_LOOP_PHOTOS][key],
                row_idx
            )
            if success:
                inserted = True
            else:
                errors.append(f"第 {row_idx} 行 (隐患编号 {key}) 闭环照片插入失败：{error}")
        
        # 设置行高
        if inserted:
            ws.row_dimensions[row_idx].height = ROW_HEIGHT_FOR_IMG
        
        row_idx += 1
    
    wb.save(output_path)
    return errors


# ============================================================================
# Word 文档处理模块
# ============================================================================

def find_table_by_title(doc, title_text: str) -> Optional[Any]:
    """
    通过标题段落查找表格
    
    Args:
        doc: python-docx Document 对象
        title_text: 标题文本
    
    Returns:
        找到的 Table 对象，未找到返回 None
    """
    for paragraph in doc.paragraphs:
        if title_text in paragraph.text:
            paragraph_element = paragraph._element
            parent = paragraph_element.getparent()
            paragraph_index = parent.index(paragraph_element)
            
            for j in range(paragraph_index + 1, len(parent)):
                element = parent[j]
                if element.tag.endswith('tbl'):
                    for table in doc.tables:
                        if table._element == element:
                            return table
            break
    
    return None


def find_column_index(table, header_text: str) -> int:
    """
    在表格中查找指定列标题的索引
    
    Args:
        table: python-docx Table 对象
        header_text: 列标题文本
    
    Returns:
        列索引，未找到返回 -1
    """
    if table.rows:
        header_row = table.rows[0]
        for i, cell in enumerate(header_row.cells):
            if header_text in cell.text:
                return i
    return -1


def apply_table_formatting(table) -> None:
    """
    为表格应用统一格式 (居中、仿宋字体、四号字、加粗)
    
    Args:
        table: python-docx Table 对象
    """
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.name = FONT_NAME
                    run.font.size = Pt(FONT_SIZE_PT)
                    run.font.bold = True
                    if hasattr(run._element, 'rPr') and run._element.rPr is not None:
                        if run._element.rPr.rFonts is not None:
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT_NAME)


def insert_image_to_word_cell(
    cell,
    img_data: bytes,
    width_inches: float = WORD_IMG_WIDTH_INCHES,
    height_inches: float = WORD_IMG_HEIGHT_INCHES
) -> bool:
    """
    向 Word 单元格插入图片
    
    Args:
        cell: python-docx Cell 对象
        img_data: 图片数据
        width_inches: 图片宽度 (英寸)
        height_inches: 图片高度 (英寸)
    
    Returns:
        是否成功插入
    """
    try:
        cell.text = ""
        img_buffer = resize_image_to_buffer(img_data)
        run = cell.paragraphs[0].add_run()
        run.add_picture(img_buffer, width=Inches(width_inches), height=Inches(height_inches))
        return True
    except Exception:
        cell.text = ""
        return False


def add_empty_row_message(table, message: str, num_columns: int) -> None:
    """
    向空表格添加说明行
    
    Args:
        table: python-docx Table 对象
        message: 说明文本
        num_columns: 表格列数
    """
    new_row = table.add_row()
    cells = new_row.cells
    if len(cells) >= num_columns:
        cells[0].text = "1"
        cells[1].text = message
        for i in range(2, len(cells)):
            cells[i].text = "/"


def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    准备 DataFrame，格式化日期列
    
    Args:
        df: 原始 DataFrame
    
    Returns:
        处理后的 DataFrame
    
    Raises:
        ValueError: 缺少必需列
    """
    missing_cols = [col for col in REQUIRED_EXCEL_COLUMNS if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Excel 文件中缺少以下列：{missing_cols}")
    
    df = df.copy()
    for date_col in ['发现时间', '要求闭环时间']:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce').dt.strftime('%Y-%m-%d')
    
    return df


def add_row_to_table_with_images(
    table,
    row: pd.Series,
    serial_number: int,
    folder_images: Dict[str, Dict[str, bytes]],
    photo_folders: List[str]
) -> None:
    """
    向表格添加一行数据并插入图片
    
    Args:
        table: python-docx Table 对象
        row: DataFrame 行数据
        serial_number: 序号
        folder_images: 图片字典
        photo_folders: 需要插入的图片文件夹列表
    """
    new_row = table.add_row()
    cells = new_row.cells
    
    # 填充序号
    cells[0].text = str(serial_number)
    
    # 填充数据列
    column_mapping = [
        ('异常事项', 1),
        ('异常类别', 2),
        ('班组', 3),
        ('整改人', 4),
        ('发现时间', 5),
        ('要求闭环时间', 6)
    ]
    
    for col_name, col_idx in column_mapping:
        if col_idx < len(cells):
            value = row.get(col_name, '')
            cells[col_idx].text = str(value) if pd.notna(value) else ""
    
    # 插入图片
    hazard_id = str(row.get('隐患编号', ''))
    
    for photo_folder in photo_folders:
        photo_col_name = "隐患照片" if photo_folder == FOLDER_HAZARD_PHOTOS else "闭环照片"
        photo_col_idx = find_column_index(table, photo_col_name)
        
        if photo_col_idx != -1 and hazard_id in folder_images[photo_folder]:
            insert_image_to_word_cell(
                cells[photo_col_idx],
                folder_images[photo_folder][hazard_id]
            )
        elif photo_col_idx != -1:
            cells[photo_col_idx].text = ""


def generate_report_core(
    excel_path: str,
    doc_template_path: str,
    zip_path: str,
    output_path: str,
    report_type: str
) -> None:
    """
    生成报告的核心逻辑
    
    Args:
        excel_path: Excel 文件路径
        doc_template_path: Word 模板路径
        zip_path: ZIP 文件路径
        output_path: 输出文件路径
        report_type: 报告类型 ('check' 或 'closure')
    """
    # 读取并准备数据
    df = pd.read_excel(excel_path)
    df = prepare_dataframe(df)
    
    # 加载模板
    doc = Document(doc_template_path)
    
    # 提取图片
    folder_images = extract_zip_to_dict(zip_path)
    
    # 查找表格
    tables_config = {
        'env_table': "二、环境保护",
        'general_table': "一、本期存在主要问题",
        'major_table': "三、重大事故隐患检查情况"
    }
    
    tables = {}
    for var_name, title in tables_config.items():
        tables[var_name] = find_table_by_title(doc, title)
    
    # 计数器
    counters = {'env_counter': 0, 'general_counter': 0, 'major_counter': 0}
    
    # 处理每一行数据
    for _, row in df.iterrows():
        category = str(row.get('异常类别', ''))
        level = str(row.get('隐患级别', ''))
        
        # 环境保护类别
        if category == '环境保护' and tables['env_table']:
            counters['env_counter'] += 1
            add_row_to_table_with_images(
                tables['env_table'], row,
                counters['env_counter'],
                folder_images,
                [FOLDER_HAZARD_PHOTOS]
            )
        
        # 一般隐患 (非环保)
        if level == '一般隐患' and category != '环境保护' and tables['general_table']:
            counters['general_counter'] += 1
            add_row_to_table_with_images(
                tables['general_table'], row,
                counters['general_counter'],
                folder_images,
                [FOLDER_HAZARD_PHOTOS]
            )
        
        # 重大隐患 (非环保)
        if level == '重大隐患' and category != '环境保护' and tables['major_table']:
            counters['major_counter'] += 1
            add_row_to_table_with_images(
                tables['major_table'], row,
                counters['major_counter'],
                folder_images,
                [FOLDER_HAZARD_PHOTOS]
            )
    
    # 处理空表格
    empty_messages = {
        'env_table': "本次检查未发现公司存在环境保护相关问题",
        'major_table': "根据《重大事故隐患清单》逐一排查，发现公司未存在重大事故隐患。"
    }
    
    for table_name, message in empty_messages.items():
        if tables[table_name] and counters[f"{table_name.split('_')[0]}_counter"] == 0:
            num_cols = 9 if report_type == 'closure' else 8
            add_empty_row_message(tables[table_name], message, num_cols)
    
    # 应用格式
    for table in tables.values():
        if table:
            apply_table_formatting(table)
    
    # 保存文档
    doc.save(output_path)


def generate_check_report(
    excel_path: str,
    doc_template_path: str,
    zip_path: str,
    output_path: str
) -> None:
    """
    生成检查报告
    
    Args:
        excel_path: Excel 文件路径
        doc_template_path: Word 模板路径
        zip_path: ZIP 文件路径
        output_path: 输出文件路径
    """
    generate_report_core(excel_path, doc_template_path, zip_path, output_path, 'check')


def generate_closure_report(
    excel_path: str,
    doc_template_path: str,
    zip_path: str,
    output_path: str
) -> None:
    """
    生成闭环报告
    
    Args:
        excel_path: Excel 文件路径
        doc_template_path: Word 模板路径
        zip_path: ZIP 文件路径
        output_path: 输出文件路径
    """
    generate_report_core(excel_path, doc_template_path, zip_path, output_path, 'closure')


# ============================================================================
# GUI 应用类
# ============================================================================

class App:
    """隐患整改台账与报告生成工具的图形界面应用"""
    
    def __init__(self, root: tk.Tk):
        """
        初始化应用
        
        Args:
            root: Tkinter 根窗口
        """
        self.root = root
        self.root.title("隐患整改台账与报告生成工具")
        
        # 文件路径存储
        self.excel_path = ""
        self.check_report_template_path = ""
        self.closure_report_template_path = ""
        self.zip_path = ""
        
        self.setup_ui()
    
    def setup_ui(self) -> None:
        """设置用户界面"""
        frame = ttk.Frame(self.root, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 文件选择区域
        file_selections = [
            ("选择隐患整改通知单", self.select_excel, "excel_label"),
            ("选择检查报告模板", self.select_check_report_template, "check_report_template_label"),
            ("选择闭环报告模板", self.select_closure_report_template, "closure_report_template_label"),
            ("选择图片压缩包", self.select_zip, "zip_label")
        ]
        
        for i, (btn_text, command, label_attr) in enumerate(file_selections):
            ttk.Button(frame, text=btn_text, command=command).grid(
                row=i, column=0, pady=5, sticky=tk.W
            )
            label = ttk.Label(frame, text="未选择")
            label.grid(row=i, column=1, padx=10, sticky=tk.W)
            setattr(self, label_attr, label)
        
        # 按钮区域
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        
        actions = [
            ("隐患整改台账生成", self.generate_excel_report),
            ("检查报告生成", self.generate_check_report),
            ("闭环报告生成", self.generate_closure_report)
        ]
        
        for btn_text, command in actions:
            ttk.Button(button_frame, text=btn_text, command=command).pack(
                side=tk.LEFT, padx=5
            )
    
    def select_excel(self) -> None:
        """选择 Excel 文件"""
        path = filedialog.askopenfilename(
            title="选择隐患整改通知单",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            self.excel_path = path
            self.excel_label.config(text=os.path.basename(path))
    
    def select_check_report_template(self) -> None:
        """选择检查报告模板"""
        path = filedialog.askopenfilename(
            title="选择检查报告模板",
            filetypes=[("Word files", "*.docx")]
        )
        if path:
            self.check_report_template_path = path
            self.check_report_template_label.config(text=os.path.basename(path))
    
    def select_closure_report_template(self) -> None:
        """选择闭环报告模板"""
        path = filedialog.askopenfilename(
            title="选择闭环报告模板",
            filetypes=[("Word files", "*.docx")]
        )
        if path:
            self.closure_report_template_path = path
            self.closure_report_template_label.config(text=os.path.basename(path))
    
    def select_zip(self) -> None:
        """选择 ZIP 文件"""
        path = filedialog.askopenfilename(
            title="选择图片压缩包",
            filetypes=[("ZIP files", "*.zip")]
        )
        if path:
            self.zip_path = path
            self.zip_label.config(text=os.path.basename(path))
    
    def generate_excel_report(self) -> None:
        """生成带图片的 Excel 台账"""
        if not self.excel_path or not self.zip_path:
            messagebox.showerror("错误", "请先选择隐患整改通知单和图片压缩包！")
            return
        
        try:
            output_path = os.path.splitext(self.excel_path)[0] + "_带图片.xlsx"
            errors = embed_images_to_excel(self.excel_path, self.zip_path, output_path)
            
            msg = f"处理完成！\n输出文件：\n{output_path}"
            
            if errors:
                msg += "\n\n⚠️ 部分图片插入失败，详见下方错误信息。"
                messagebox.showinfo("处理完成（含警告）", msg)
                
                error_text = "\n".join(errors[:20])
                if len(errors) > 20:
                    error_text += f"\n... 还有 {len(errors) - 20} 条错误未显示"
                messagebox.showwarning("插入失败详情", error_text)
            else:
                messagebox.showinfo("成功", msg)
        
        except Exception as e:
            messagebox.showerror("严重错误", f"程序运行失败：\n{str(e)}")
    
    def generate_check_report(self) -> None:
        """生成检查报告"""
        if not all([self.excel_path, self.check_report_template_path, self.zip_path]):
            messagebox.showerror("错误", "请先选择隐患整改通知单、检查报告模板和图片压缩包！")
            return
        
        try:
            output_path = os.path.splitext(self.check_report_template_path)[0] + "_检查报告.docx"
            generate_check_report(
                self.excel_path,
                self.check_report_template_path,
                self.zip_path,
                output_path
            )
            messagebox.showinfo("成功", f"检查报告生成完成！\n输出文件：\n{output_path}")
        
        except Exception as e:
            messagebox.showerror("严重错误", f"程序运行失败：\n{str(e)}")
    
    def generate_closure_report(self) -> None:
        """生成闭环报告"""
        if not all([self.excel_path, self.closure_report_template_path, self.zip_path]):
            messagebox.showerror("错误", "请先选择隐患整改通知单、闭环报告模板和图片压缩包！")
            return
        
        try:
            output_path = os.path.splitext(self.closure_report_template_path)[0] + "_闭环报告.docx"
            generate_closure_report(
                self.excel_path,
                self.closure_report_template_path,
                self.zip_path,
                output_path
            )
            messagebox.showinfo("成功", f"闭环报告生成完成！\n输出文件：\n{output_path}")
        
        except Exception as e:
            messagebox.showerror("严重错误", f"程序运行失败：\n{str(e)}")


# ============================================================================
# 主程序入口
# ============================================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
