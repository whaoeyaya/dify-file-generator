from pptx import Presentation
from pptx.util import Pt
import os

# 自动定位模板路径
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(CURRENT_DIR, 'template.pptx')

def generate_ppt_from_template(ppt_data, basic_info, output_path):
    # 载入现有模板，而不是创建空白 Presentation()
    prs = Presentation(TEMPLATE_PATH if os.path.exists(TEMPLATE_PATH) else None)
    
    # 1. 替换/设置首页标题（假设 Slide 1 是封面）
    if len(prs.slides) > 0:
        slide_layout = prs.slides[0]
        for shape in slide_layout.shapes:
            if shape.has_text_frame:
                # 匹配封面标题占位符或文本框
                if "课题" in shape.text_frame.text or shape.text_frame.text == "":
                    shape.text_frame.text = basic_info.get("lesson_title", "教学课件")
                    break

    # 2. 动态生成内容页（使用模板中的“标题+内容” Layout，通常是 Index 1）
    bullet_slide_layout = prs.slide_layouts[1] 
    
    for slide_info in ppt_data.get("slides", []):
        slide = prs.slides.add_slide(bullet_slide_layout)
        
        # 填充标题
        if slide.shapes.title:
            slide.shapes.title.text = slide_info.get("title", "教学环节")
        
        # 填充正文 Bullet Points
        # 寻找正文文本框 (Body Placeholder)
        for shape in slide.shapes:
            if shape.has_text_frame and shape != slide.shapes.title:
                tf = shape.text_frame
                tf.clear() # 清除默认占位文本
                
                for idx, pt_text in enumerate(slide_info.get("bullet_points", [])):
                    p = tf.add_paragraph() if idx > 0 else tf.paragraphs[0]
                    p.text = pt_text
                    p.font.size = Pt(18)
                break

    prs.save(output_path)
    return output_path

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    """设置单元格背景颜色"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=140, bottom=140, left=150, right=150):
    """设置单元格边距 (dxa 80 = 4pt)"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin_name, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{margin_name}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def generate_word_lesson_plan(word_data, basic_info, output_path):
    doc = Document()
    
    # 页面边距设置 (1英寸)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # 1. 大标题
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(f"《{basic_info.get('lesson_title', '教学设计')}》整体教学设计")
    run.font.name = '黑体'
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D) # 深蓝标题

    doc.add_paragraph()

    # 2. 基础信息表 (2行4列)
    info_table = doc.add_table(rows=2, cols=4)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.style = 'Table Grid'
    
    headers = [
        ("科 目", basic_info.get("subject", "")),
        ("年 级", basic_info.get("grade", "")),
        ("课 题", basic_info.get("lesson_title", "")),
        ("课 时", basic_info.get("period", "1课时"))
    ]
    
    cells = info_table.flat_cells
    for i, (label, val) in enumerate(headers):
        cells[i*2].text = label
        set_cell_background(cells[i*2], "F2F2F2")
        cells[i*2].paragraphs[0].runs[0].font.bold = True
        cells[i*2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        cells[i*2+1].text = val
        cells[i*2+1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_cell_margins(cells[i*2+1])

    doc.add_paragraph()

    # 3. 核心教学过程四栏表 (环节 | 教师活动 | 学生活动 | 设计意图)
    process_list = word_data.get('teaching_process', [])
    main_table = doc.add_table(rows=1 + len(process_list), cols=4)
    main_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    main_table.style = 'Table Grid'
    
    col_widths = [Inches(1.2), Inches(2.3), Inches(2.3), Inches(1.2)]
    
    # 表头设置
    hdr_titles = ["教学环节", "教师活动", "学生活动", "设计意图"]
    hdr_cells = main_table.rows[0].cells
    for i, title in enumerate(hdr_titles):
        hdr_cells[i].text = title
        hdr_cells[i].width = col_widths[i]
        set_cell_background(hdr_cells[i], "1F497D") # 表头深蓝背景
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if p.runs:
            p.runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            p.runs[0].font.bold = True

    # 逐行填入教学过程
    for row_idx, stage_data in enumerate(process_list, start=1):
        row_cells = main_table.rows[row_idx].cells
        data_items = [
            stage_data.get("stage", ""),
            stage_data.get("teacher_activity", ""),
            stage_data.get("student_activity", ""),
            stage_data.get("design_intent", "")
        ]
        
        for c_idx, text in enumerate(data_items):
            cell = row_cells[c_idx]
            cell.width = col_widths[c_idx]
            cell.text = text
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell)
            
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(4)
            
            if c_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if p.runs:
                    p.runs[0].font.bold = True
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    doc.save(output_path)
    return output_path

# 在 SimpleHandler 的 do_POST 方法中
def do_POST(self):
    content_length = int(self.headers['Content-Length'])
    post_data = self.rfile.read(content_length)
    req_json = json.loads(post_data.decode('utf-8'))

    # 提取大模型返回的结构化 JSON
    model_output = req_json.get('model_output', {})
    basic_info = model_output.get('basic_info', {})
    ppt_data = model_output.get('ppt_data', {})
    word_data = model_output.get('word_data', {})

    lesson_title = basic_info.get('lesson_title', 'default')
    
    # 动态生成唯一文件名
    ppt_filename = f"{lesson_title}_教学课件.pptx"
    word_filename = f"{lesson_title}_教学设计.docx"
    
    ppt_path = os.path.join(DOWNLOAD_DIR, ppt_filename)
    word_path = os.path.join(DOWNLOAD_DIR, word_filename)

    # 1. 动态填充 PPT 模板
    generate_ppt_from_template(ppt_data, basic_info, ppt_path)
    
    # 2. 动态生成 Word 教案表格
    generate_word_lesson_plan(word_data, basic_info, word_path)

    # 返回给 Dify 的下载链接或文件列表
    response_data = {
        "status": "success",
        "ppt_url": f"/static/{ppt_filename}",
        "word_url": f"/static/{word_filename}"
    }
    
    self.send_response(200)
    self.send_header('Content-Type', 'application/json')
    self.end_headers()
    self.wfile.write(json.dumps(response_data).encode('utf-8'))
