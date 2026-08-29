import os
import json
import urllib.parse
import mimetypes
from http.server import HTTPServer, BaseHTTPRequestHandler
from pptx import Presentation
from pptx.util import Pt
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
# ---------------------------------------------------------------------------
def generate_ppt_from_template(ppt_data, basic_info, output_path):
    prs = Presentation(PPT_TEMPLATE_PATH if os.path.exists(PPT_TEMPLATE_PATH) else None)
    
    # 设置封面标题
    if len(prs.slides) > 0:
        slide_layout = prs.slides[0]
        for shape in slide_layout.shapes:
            if shape.has_text_frame:
                if "课题" in shape.text_frame.text or shape.text_frame.text == "":
                    shape.text_frame.text = basic_info.get("lesson_title", "教学课件")
                    break

    # 动态生成内容页
    bullet_slide_layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]
    
    for slide_info in ppt_data.get("slides", []):
        slide = prs.slides.add_slide(bullet_slide_layout)
        if slide.shapes.title:
            slide.shapes.title.text = slide_info.get("title", "教学环节")
        
        for shape in slide.shapes:
            if shape.has_text_frame and shape != slide.shapes.title:
                tf = shape.text_frame
                tf.clear()
                for idx, pt_text in enumerate(slide_info.get("bullet_points", [])):
                    p = tf.add_paragraph() if idx > 0 else tf.paragraphs[0]
                    p.text = pt_text
                    p.font.size = Pt(18)
                break

    prs.save(output_path)
    return output_path

# ---------------------------------------------------------------------------
# 2. Word 教学设计生成逻辑 (基于现有模板填充)
# ---------------------------------------------------------------------------

def replace_placeholder_in_paragraph(p, mapping):
    """替换段落中的占位符，同时保留字体格式"""
    for key, value in mapping.items():
        if key in p.text:
            p.text = p.text.replace(key, str(value))

def replace_placeholder_in_table(table, mapping):
    """扫描表格所有单元格，精准替换占位符"""
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for key, value in mapping.items():
                    if key in p.text:
                        p.text = p.text.replace(key, str(value))

def generate_word_lesson_plan(word_data, basic_info, output_path):
    try:
        if os.path.exists(WORD_TEMPLATE_PATH):
            doc = Document(WORD_TEMPLATE_PATH)
        else:
            doc = Document()

        # 1. 整理教学目标格式
        obj = word_data.get("teaching_objectives", {})
        if isinstance(obj, dict):
            obj_str = f"1. 知识与技能：{obj.get('knowledge', '')}\n2. 过程与方法：{obj.get('ability', '')}\n3. 情感态度与价值观：{obj.get('literacy', '')}"
        else:
            obj_str = str(obj)

        # 2. 整理重难点格式
        kp = word_data.get('key_points', '')
        dp = word_data.get('difficult_points', '')
        key_and_difficult = f"教学重点：{kp}\n教学难点：{dp}"

        # 3. 整理教法与学法格式
        methods = word_data.get("teaching_methods", {})
        if isinstance(methods, dict):
            methods_str = f"教法：{methods.get('teacher_method', '情境教学法')}\n学法：{methods.get('student_method', '自主探究、合作交流')}"
        else:
            methods_str = str(methods) if str(methods) else "教法：情境教学法\n学法：自主探究、合作交流"

        # 4. 整理归纳总结与作业设计格式
        summary = word_data.get("summary_and_homework", {})
        if isinstance(summary, dict):
            sum_str = f"【归纳总结】：{summary.get('summary', '')}\n【作业设计】：{summary.get('homework', '')}\n【素养发展】：{summary.get('literacy_development', '')}"
        else:
            sum_str = str(summary) if str(summary) else "归纳总结：梳理小数加减法算理。\n作业设计：结合本土超市购物清单计算总价。"

        # 5. 整理五环节教学流程数据
        process_list = word_data.get('teaching_process', [])
        p1 = process_list[0] if len(process_list) > 0 else {}
        p2 = process_list[1] if len(process_list) > 1 else {}
        p3 = process_list[2] if len(process_list) > 2 else {}
        p4 = process_list[3] if len(process_list) > 3 else {}
        p5 = process_list[4] if len(process_list) > 4 else {}

        # 6. 构建占位符映射字典 (Mapping)
        mapping = {
            # 基础信息
            "{{teacher_name}}": basic_info.get("teacher_name", "教师"),
            "{{teacher_unit}}": basic_info.get("teacher_unit", "当地实验小学"),
            "{{lesson_date}}": basic_info.get("lesson_date", "2026年"),
            "{{stage}}": basic_info.get("stage", "小学"),
            "{{subject}}": basic_info.get("subject", "数学"),
            "{{grade}}": basic_info.get("grade", "四年级"),
            "{{lesson_time}}": basic_info.get("lesson_time", "40分钟"),
            "{{lesson_type}}": basic_info.get("lesson_type", "新授课"),
            "{{lesson_title}}": basic_info.get('lesson_title', '教学设计'),

            # 教学分析与目标
            "{{textbook_analysis}}": basic_info.get("textbook_analysis", word_data.get("textbook_analysis", "结合本土实际素材深入浅出阐述概念。")),
            "{{student_analysis}}": basic_info.get("student_analysis", word_data.get("student_analysis", "学生具备基础计算能力，但需加强生活应用。")),
            "{{teaching_objectives}}": obj_str,
            "{{key_and_difficult_points}}": key_and_difficult,
            "{{teaching_methods}}": methods_str,

            # 教学流程五环节
            "{{t1}}": p1.get("teacher_activity", ""), "{{s1}}": p1.get("student_activity", ""), "{{d1}}": p1.get("design_intent", ""),
            "{{t2}}": p2.get("teacher_activity", ""), "{{s2}}": p2.get("student_activity", ""), "{{d2}}": p2.get("design_intent", ""),
            "{{t3}}": p3.get("teacher_activity", ""), "{{s3}}": p3.get("student_activity", ""), "{{d3}}": p3.get("design_intent", ""),
            "{{t4}}": p4.get("teacher_activity", ""), "{{s4}}": p4.get("student_activity", ""), "{{d4}}": p4.get("design_intent", ""),
            "{{t5}}": p5.get("teacher_activity", ""), "{{s5}}": p5.get("student_activity", ""), "{{d5}}": p5.get("design_intent", ""),

            # 尾部模块
            "{{summary_and_homework}}": sum_str,
            "{{board_design}}": word_data.get("board_design", "板书设计：核心概念与算理推导"),
            "{{reflection}}": word_data.get("reflection", "结合本土真实生活情境，有效激发了学生学习兴趣。")
        }

        # 7. 执行全局替换（标题段落 + 表格区域）
        for p in doc.paragraphs:
            replace_placeholder_in_paragraph(p, mapping)

        for table in doc.tables:
            replace_placeholder_in_table(table, mapping)

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
