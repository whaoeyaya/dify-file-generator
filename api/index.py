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
def generate_word_lesson_plan(word_data, basic_info, output_path):
    # 如果有 template.docx 模板，则载入模板；否则新建文档
    if os.path.exists(WORD_TEMPLATE_PATH):
        doc = Document(WORD_TEMPLATE_PATH)
    else:
        doc = Document()

    lesson_title = basic_info.get('lesson_title', '教学设计')

    # A. 替换文档主标题
    for p in doc.paragraphs:
        if "教学设计" in p.text:
            p.text = f"《{lesson_title}》教学设计"
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if len(p.runs) > 0:
                p.runs[0].font.name = '黑体'
                p.runs[0].font.size = Pt(18)
                p.runs[0].font.bold = True
            break

    # B. 填充主表格数据
    if len(doc.tables) > 0:
        table = doc.tables[0] # 获取主教案大表
        
        # 逐行遍历表格进行关键字匹配填充
        for row in table.rows:
            row_text = "".join([cell.text for cell in row.cells])
            
            # 1. 基础信息
            if "课题名称" in row_text and len(row.cells) >= 2:
                row.cells[1].text = lesson_title
            elif "教材及版本" in row_text and len(row.cells) >= 2:
                row.cells[1].text = basic_info.get("textbook", "标准教材")
            elif "授课教师" in row_text and len(row.cells) >= 2:
                row.cells[1].text = basic_info.get("teacher_name", "教师")
            elif "学段" in row_text:
                for c_idx, cell in enumerate(row.cells):
                    if "学科" in cell.text and c_idx + 1 < len(row.cells):
                        row.cells[c_idx+1].text = basic_info.get("subject", "数学")
                    if "适用年级" in cell.text and c_idx + 1 < len(row.cells):
                        row.cells[c_idx+1].text = basic_info.get("grade", "通用年级")

            # 2. 教学目标
            elif "教学目标" in row.cells[0].text:
                objectives = word_data.get("teaching_objectives", "")
                if isinstance(objectives, dict):
                    obj_text = f"1. 知识与技能：{objectives.get('knowledge', '')}\n2. 过程与方法：{objectives.get('ability', '')}\n3. 情感态度与价值观：{objectives.get('literacy', '')}"
                else:
                    obj_text = str(objectives)
                row.cells[1].text = obj_text

            # 3. 教学重难点
            elif "教学重难点" in row.cells[0].text:
                kp = word_data.get("key_points", "")
                dp = word_data.get("difficult_points", "")
                row.cells[1].text = f"教学重点：{kp}\n教学难点：{dp}"

            # 4. 板书设计
            elif "板书设计" in row.cells[0].text:
                row.cells[1].text = word_data.get("board_design", f"板书设计：{lesson_title}\n1. 核心概念与推导\n2. 练习与总结")

            # 5. 分层作业
            elif "分层作业" in row.cells[0].text:
                homework = word_data.get("homework", {})
                if isinstance(homework, dict):
                    hw_text = f"基础巩固：{homework.get('basic', '')}\n拓展提升：{homework.get('advanced', '')}"
                else:
                    hw_text = str(homework)
                row.cells[1].text = hw_text

            # 6. 教学反思
            elif "课后反思" in row.cells[0].text or "教学反思" in row.cells[0].text:
                row.cells[1].text = word_data.get("reflection", "根据课堂实际生成反馈，注重本土素材与核心概念的深度结合。")

    # C. 动态填充【五环节教学流程】明细行
    process_list = word_data.get('teaching_process', [])
    if len(doc.tables) > 0 and len(process_list) > 0:
        table = doc.tables[0]
        # 寻找“教学环节”所在的表头行索引
        header_row_idx = -1
        for idx, row in enumerate(table.rows):
            if "教学环节" in row.cells[0].text and "教师活动" in row.cells[1].text:
                header_row_idx = idx
                break
        
        if header_row_idx != -1:
            # 将具体环节写入随后的数据行中
            for i, process in enumerate(process_list):
                target_row_idx = header_row_idx + 1 + i
                if target_row_idx < len(table.rows):
                    row = table.rows[target_row_idx]
                    if len(row.cells) >= 4:
                        row.cells[0].text = process.get("stage", "")
                        row.cells[1].text = process.get("teacher_activity", "")
                        row.cells[2].text = process.get("student_activity", "")
                        row.cells[3].text = process.get("design_intent", "")

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
