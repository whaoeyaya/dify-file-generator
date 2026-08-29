import os
import json
import urllib.parse
import mimetypes
from http.server import HTTPServer, BaseHTTPRequestHandler
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(CURRENT_DIR, 'static')
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

PPT_TEMPLATE_PATH = os.path.join(CURRENT_DIR, 'template.pptx')
WORD_TEMPLATE_PATH = os.path.join(CURRENT_DIR, 'template.docx')

# ---------------------------------------------------------------------------
# 1. PPT 生成逻辑

import os
import json
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

PPT_TEMPLATE_NAME = "template.pptx"

def safe_json_parse(data):
    """确保数据转为 dict 格式"""
    if isinstance(data, dict):
        return data
    if isinstance(data, str) and data.strip():
        try:
            return json.loads(data)
        except Exception:
            return {}
    return {}

def clear_existing_slides(prs, n=2):
    """在生成新内容前，先彻底清理掉模板自带的前 N 页"""
    slide_ids = [slide.slide_id for slide in prs.slides]
    delete_count = min(n, len(slide_ids))
    for _ in range(delete_count):
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

def set_font_style(run, size_pt, bold=False, color_rgb=(51, 51, 51)):
    """统一设置字体大小与颜色"""
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*color_rgb)

def generate_ppt_lesson_plan(ppt_data, output_path):
    try:
        ppt_data = safe_json_parse(ppt_data)
        basic_info = ppt_data.get('basic_info', {})
        word_data = ppt_data.get('word_data', {})
        
        current_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
        template_path = os.path.join(current_dir, PPT_TEMPLATE_NAME)
        
        # 1. 加载模板
        if os.path.exists(template_path):
            prs = Presentation(template_path)
            # ★【核心修改 1】：优先在最开始就删掉前两页模板，防止误删新生成的内容
            clear_existing_slides(prs, n=2)
        else:
            prs = Presentation()

        # 使用空白版式（Layout 6 通常是纯白版式，防止被模板背景遮挡）
        blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[0]

        # ---------------- 2. 开始生成新页面 ----------------
        
        # Slide 1: 封面页
        slide_cover = prs.slides.add_slide(blank_layout)
        title_box = slide_cover.shapes.add_textbox(Inches(1), Inches(1.5), Inches(11.3), Inches(4.5))
        tf = title_box.text_frame
        tf.word_wrap = True
        
        lesson_title = basic_info.get('lesson_title') or word_data.get('lesson_title', '教学设计')
        p1 = tf.paragraphs[0]
        p1.text = f"《{lesson_title}》教学设计"
        set_font_style(p1.runs[0], size_pt=40, bold=True, color_rgb=(0, 51, 102))
        
        p2 = tf.add_paragraph()
        sub_info = f"\n学科：{basic_info.get('subject', '数学')}  |  年级：{basic_info.get('grade', '')}\n执教教师：{basic_info.get('teacher_name', '')}"
        p2.text = sub_info
        set_font_style(p2.runs[0], size_pt=24, color_rgb=(80, 80, 80))

        # Slide 2: 教学目标与重难点
        slide_target = prs.slides.add_slide(blank_layout)
        t_box = slide_target.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.5), Inches(5.8))
        tf2 = t_box.text_frame
        tf2.word_wrap = True
        
        p_t = tf2.paragraphs[0]
        p_t.text = "教学目标与重难点"
        set_font_style(p_t.runs[0], size_pt=32, bold=True, color_rgb=(0, 51, 102))
        
        objs = word_data.get('teaching_objectives', {})
        target_text = (
            f"\n【教学重点】\n{word_data.get('key_points', '')}\n\n"
            f"【教学难点】\n{word_data.get('difficult_points', '')}\n\n"
            f"【核心目标】\n"
            f"• 知识与技能：{objs.get('knowledge', '')}\n"
            f"• 过程与方法：{objs.get('ability', '')}"
        )
        p_c = tf2.add_paragraph()
        p_c.text = target_text
        set_font_style(p_c.runs[0], size_pt=20)

        # Slide 3-7: 5 个教学环节
        process_list = word_data.get('teaching_process', [])
        for item in process_list:
            if not isinstance(item, dict): continue
            
            slide_p = prs.slides.add_slide(blank_layout)
            p_box = slide_p.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.5), Inches(6.2))
            tf_p = p_box.text_frame
            tf_p.word_wrap = True
            
            # 环节标题
            p_stage = tf_p.paragraphs[0]
            p_stage.text = f"教学环节：{item.get('stage', '')}"
            set_font_style(p_stage.runs[0], size_pt=28, bold=True, color_rgb=(0, 51, 102))
            
            # 活动与意图
            p_act = tf_p.add_paragraph()
            p_act.text = (
                f"\n【教师活动】\n{item.get('teacher_activity', '')}\n\n"
                f"【学生活动】\n{item.get('student_activity', '')}\n\n"
                f"【设计意图】\n{item.get('design_intent', '')}"
            )
            set_font_style(p_act.runs[0], size_pt=18)

        # 3. 保存 PPT
        prs.save(output_path)
        return output_path

    except Exception as e:
        print(f"[ERROR PPT Generation] {str(e)}")
        return None
# ---------------------------------------------------------------------------

import os
import json
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

def safe_json_parse(data):
    """确保数据转为 dict 格式"""
    if isinstance(data, dict):
        return data
    if isinstance(data, str) and data.strip():
        try:
            return json.loads(data)
        except Exception:
            return {}
    return {}

def generate_word_lesson_plan(word_data, basic_info, output_path):
    try:
        # 1. 动态安全解析 JSON 数据
        basic_info = safe_json_parse(basic_info)
        word_data = safe_json_parse(word_data)

        # 2. 读取 Word 模板
        if os.path.exists(WORD_TEMPLATE_PATH):
            doc = Document(WORD_TEMPLATE_PATH)
        else:
            doc = Document()

        lesson_title = basic_info.get('lesson_title') or word_data.get('lesson_title', '')

        # A. 替换主标题
        for p in doc.paragraphs:
            if "教学设计" in p.text:
                p.text = f"《{lesson_title}》教学设计" if lesson_title else p.text
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                break

        # B. 填充主表格
        if len(doc.tables) > 0:
            table = doc.tables[0]

            for idx, row in enumerate(table.rows):
                row_text = "".join([c.text.strip() for c in row.cells])

                # 1. 基础信息行（完全按输入动态填写）
                if any(k in row_text for k in ["授课教师", "学段", "授课时间"]):
                    for c_idx, cell in enumerate(row.cells):
                        cell_txt = cell.text.strip()
                        if cell_txt == "授课教师" and c_idx + 1 < len(row.cells):
                            row.cells[c_idx + 1].text = str(basic_info.get("teacher_name", ""))
                        elif cell_txt == "授课教师单位" and c_idx + 1 < len(row.cells):
                            row.cells[c_idx + 1].text = str(basic_info.get("teacher_unit", ""))
                        elif cell_txt == "授课日期" and c_idx + 1 < len(row.cells):
                            row.cells[c_idx + 1].text = str(basic_info.get("lesson_date", ""))
                        elif cell_txt == "学段" and c_idx + 1 < len(row.cells):
                            row.cells[c_idx + 1].text = str(basic_info.get("stage", ""))
                        elif cell_txt == "学科" and c_idx + 1 < len(row.cells):
                            row.cells[c_idx + 1].text = str(basic_info.get("subject", ""))
                        elif cell_txt == "适用年级" and c_idx + 1 < len(row.cells):
                            row.cells[c_idx + 1].text = str(basic_info.get("grade", ""))
                        elif cell_txt == "授课时间" and c_idx + 1 < len(row.cells):
                            row.cells[c_idx + 1].text = str(basic_info.get("lesson_time", ""))
                        elif cell_txt == "课型" and c_idx + 1 < len(row.cells):
                            row.cells[c_idx + 1].text = str(basic_info.get("lesson_type", ""))

                # 2. 文本区域填充
                elif row.cells[0].text.strip().startswith("课题"):
                    row.cells[-1].text = lesson_title

                elif "教材分析" in row.cells[0].text:
                    row.cells[-1].text = str(basic_info.get("textbook_analysis") or word_data.get("textbook_analysis", ""))

                elif "学情分析" in row.cells[0].text:
                    row.cells[-1].text = str(basic_info.get("student_analysis") or word_data.get("student_analysis", ""))

                elif "教学目标" in row.cells[0].text:
                    obj = word_data.get("teaching_objectives") or basic_info.get("teaching_objectives", {})
                    if isinstance(obj, dict):
                        k = obj.get('knowledge') or obj.get('knowledge_and_skills', '')
                        a = obj.get('ability') or obj.get('process_and_methods', '')
                        l = obj.get('literacy') or obj.get('emotions_and_values', '')
                        row.cells[-1].text = f"1. 知识与技能：{k}\n2. 过程与方法：{a}\n3. 情感态度与价值观：{l}"
                    else:
                        row.cells[-1].text = str(obj)

                elif "教学重难点" in row.cells[0].text:
                    kp = word_data.get('key_points', '')
                    dp = word_data.get('difficult_points', '')
                    row.cells[-1].text = f"教学重点：{kp}\n教学难点：{dp}"

                elif "教法与学法" in row.cells[0].text:
                    tm = word_data.get("teaching_methods", "")
                    if isinstance(tm, dict):
                        row.cells[-1].text = f"教法：{tm.get('teacher_method', '')}\n学法：{tm.get('student_method', '')}"
                    else:
                        row.cells[-1].text = str(tm)

                elif "归纳总结" in row_text or "作业设计" in row_text:
                    sh = word_data.get("summary_and_homework") or word_data.get("homework", {})
                    if isinstance(sh, dict):
                        row.cells[-1].text = f"【归纳总结】：{sh.get('summary', '')}\n【作业设计】：{sh.get('homework', '')}\n【素养发展】：{sh.get('literacy_development', '')}"
                    else:
                        row.cells[-1].text = str(sh)

                elif "板书设计" in row_text and idx + 1 < len(table.rows):
                    table.rows[idx + 1].cells[-1].text = str(word_data.get("board_design", ""))

                elif "课后反思" in row_text and idx + 1 < len(table.rows):
                    table.rows[idx + 1].cells[-1].text = str(word_data.get("reflection", ""))

                # 3. 五环节自动化填充算法（自动匹配 5 个环节）
                else:
                    first_cell = row.cells[0].text.strip()
                    process_list = word_data.get('teaching_process', [])

                    stage_keywords = {
                        "创设": ["创设", "导入", "情境"],
                        "探究": ["探究", "建构", "新知"],
                        "理解": ["理解", "应用", "巩固"],
                        "迁移": ["迁移", "拓展", "延伸"],
                        "创新": ["创新", "提升", "综合"]
                    }

                    if isinstance(process_list, list):
                        for item in process_list:
                            if not isinstance(item, dict): continue
                            p_stage = str(item.get("stage", ""))
                            
                            # 寻找匹配的环节
                            for key, kw_list in stage_keywords.items():
                                if any(kw in first_cell for kw in kw_list) and any(kw in p_stage for kw in kw_list):
                                    if len(row.cells) >= 4:
                                        row.cells[1].text = str(item.get("teacher_activity", ""))
                                        row.cells[2].text = str(item.get("student_activity", ""))
                                        row.cells[3].text = str(item.get("design_intent", ""))
                                    break

        doc.save(output_path)
    except Exception as e:
        print(f"[ERROR Word Generation] {str(e)}")
        doc = Document()
        doc.add_heading(f"《{basic_info.get('lesson_title', '教学设计')}》", 0)
        doc.add_paragraph(str(word_data))
        doc.save(output_path)

    return output_path

# ---------------------------------------------------------------------------
# 3. HTTP 服务 Handler
# ---------------------------------------------------------------------------
class SimpleHandler(BaseHTTPRequestHandler):

    def do_GET(self):
        """处理文件下载"""
        parsed_path = urllib.parse.urlparse(self.path).path
        if parsed_path.startswith('/static/'):
            # 1. 提取文件名并进行 URL 解码 (将 %E5%A4%A7%E7%BA%B2 解转回真实中文)
            raw_filename = parsed_path.replace('/static/', '')
            decoded_filename = urllib.parse.unquote(raw_filename)
            
            file_path = os.path.join(DOWNLOAD_DIR, decoded_filename)
            
            # 调试日志：方便在 Render Logs 查看实际找的是什么文件
            print(f"[DEBUG GET] Request file: {decoded_filename}")
            print(f"[DEBUG GET] Full path: {file_path}")
            print(f"[DEBUG GET] File exists? {os.path.exists(file_path)}")

            if os.path.exists(file_path) and os.path.isfile(file_path):
                self.send_response(200)
                mime_type, _ = mimetypes.guess_type(file_path)
                self.send_header('Content-Type', mime_type or 'application/octet-stream')
                # 兼容中文下载名的 Header 标准写法
                self.send_header('Content-Disposition', f"attachment; filename*=UTF-8''{urllib.parse.quote(decoded_filename)}")
                self.send_header('Content-Length', str(os.path.getsize(file_path)))
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                with open(file_path, 'rb') as f:
                    self.wfile.write(f.read())
                return

        # 找不到文件时返回 404
        self.send_response(404)
        self.send_header('Content-Type', 'text/plain; charset=utf-8')
        self.end_headers()
        self.wfile.write(b"File Not Found")

    def do_POST(self):
        """接收 Dify 请求并生成文档"""
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length)
            req_json = json.loads(post_data.decode('utf-8'))

            model_output = req_json.get('model_output', {})
            if isinstance(model_output, str):
                clean_str = model_output.replace("```json", "").replace("```", "").strip()
                model_output = json.loads(clean_str)

            basic_info = model_output.get('basic_info', {})
            ppt_data = model_output.get('ppt_data', {})
            word_data = model_output.get('word_data', {})

            # 过滤掉文件名中的非法字符 (如 《》 / \ : * ? " < > |)
            raw_title = req_json.get('lesson_title') or basic_info.get('lesson_title') or '教学设计'
            clean_title = "".join([c for c in raw_title if c not in r'\/:*?"<>|《》'])

            ppt_filename = f"{clean_title}_教学课件.pptx"
            word_filename = f"{clean_title}_教学设计.docx"
            
            ppt_path = os.path.join(DOWNLOAD_DIR, ppt_filename)
            word_path = os.path.join(DOWNLOAD_DIR, word_filename)

            # 生成文件
            generate_ppt_from_template(ppt_data, basic_info, ppt_path)
            generate_word_lesson_plan(word_data, basic_info, word_path)

            print(f"[DEBUG POST] Generated Word at: {word_path}, exists: {os.path.exists(word_path)}")
            print(f"[DEBUG POST] Generated PPT at: {ppt_path}, exists: {os.path.exists(ppt_path)}")

            domain = "https://dify-file-generator.onrender.com"
            response_data = {
                "status": "success",
                "ppt_url": f"{domain}/static/{urllib.parse.quote(ppt_filename)}",
                "word_url": f"{domain}/static/{urllib.parse.quote(word_filename)}"
            }
            
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps(response_data).encode('utf-8'))

        except Exception as e:
            print(f"[ERROR POST] {str(e)}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"status": "error", "message": str(e)}).encode('utf-8'))

# ---------------------------------------------------------------------------
# 4. 启动 Server 逻辑
# ---------------------------------------------------------------------------
def run(server_class=HTTPServer, handler_class=SimpleHandler, port=10000):
    server_address = ('', port)
    httpd = server_class(server_address, handler_class)
    print(f"Server starting on port {port}...")
    httpd.serve_forever()

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    run(port=port)
