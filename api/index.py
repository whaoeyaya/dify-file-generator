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
    try:
        if os.path.exists(WORD_TEMPLATE_PATH):
            doc = Document(WORD_TEMPLATE_PATH)
        else:
            doc = Document()

        lesson_title = basic_info.get('lesson_title', '教学设计')

        # A. 替换主标题
        for p in doc.paragraphs:
            if "教学设计" in p.text:
                p.text = f"《{lesson_title}》教学设计"
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                break

        # B. 填充主表格
        if len(doc.tables) > 0:
            table = doc.tables[0]
            
            # 建立教学环节阶段映射词典，用于精确匹配模板左侧的“创设情境”等行
            stage_map = {
                "创设": "创设情境",
                "探究": "探究归纳建构新知",
                "理解": "知识的理解应用",
                "迁移": "知识的迁移应用",
                "创新": "知识的创新"
            }

            for idx, row in enumerate(table.rows):
                # 拼接整行非重复文本，用于鲁棒性极高的关键字识别
                row_raw_text = "".join([c.text.strip() for c in row.cells])
                
                # -------------------------------------------------------------
                # 1. 基础信息行（利用键值对在相邻 cell 的特征）
                # -------------------------------------------------------------
                if "授课教师" in row_raw_text or "学段" in row_raw_text or "授课时间" in row_raw_text:
                    for c_idx, cell in enumerate(row.cells):
                        c_txt = cell.text.strip()
                        if c_txt == "授课教师" and c_idx + 1 < len(row.cells):
                            row.cells[c_idx + 1].text = basic_info.get("teacher_name", "教师")
                        elif c_txt == "授课教师单位" and c_idx + 1 < len(row.cells):
                            row.cells[c_idx + 1].text = basic_info.get("teacher_unit", "当地实验小学")
                        elif c_txt == "授课日期" and c_idx + 1 < len(row.cells):
                            row.cells[c_idx + 1].text = basic_info.get("lesson_date", "2026年")
                        elif c_txt == "学段" and c_idx + 1 < len(row.cells):
                            row.cells[c_idx + 1].text = basic_info.get("stage", "小学")
                        elif c_txt == "学科" and c_idx + 1 < len(row.cells):
                            row.cells[c_idx + 1].text = basic_info.get("subject", "数学")
                        elif c_txt == "适用年级" and c_idx + 1 < len(row.cells):
                            row.cells[c_idx + 1].text = basic_info.get("grade", "四年级")
                        elif c_txt == "授课时间" and c_idx + 1 < len(row.cells):
                            row.cells[c_idx + 1].text = basic_info.get("lesson_time", "40分钟")
                        elif c_txt == "课型" and c_idx + 1 < len(row.cells):
                            row.cells[c_idx + 1].text = basic_info.get("lesson_type", "新授课")

                # -------------------------------------------------------------
                # 2. 单独的大块文本行（适配合并单元格：取最后一个 cell 填写）
                # -------------------------------------------------------------
                elif row.cells[0].text.strip().startswith("课题"):
                    row.cells[-1].text = lesson_title

                elif "教材分析" in row.cells[0].text:
                    row.cells[-1].text = basic_info.get("textbook_analysis", word_data.get("textbook_analysis", "结合本土实际素材深入浅出阐述概念。"))

                elif "学情分析" in row.cells[0].text:
                    row.cells[-1].text = basic_info.get("student_analysis", word_data.get("student_analysis", "学生具备基础计算能力，但需加强生活应用。"))

                elif "教学目标" in row.cells[0].text:
                    obj = word_data.get("teaching_objectives", {})
                    if isinstance(obj, dict):
                        row.cells[-1].text = f"1. 知识与技能：{obj.get('knowledge', '')}\n2. 过程与方法：{obj.get('ability', '')}\n3. 情感态度与价值观：{obj.get('literacy', '')}"
                    else:
                        row.cells[-1].text = str(obj)

                elif "教学重难点" in row.cells[0].text:
                    row.cells[-1].text = f"教学重点：{word_data.get('key_points', '')}\n教学难点：{word_data.get('difficult_points', '')}"

                elif "教法与学法" in row.cells[0].text:
                    methods = word_data.get("teaching_methods", {})
                    if isinstance(methods, dict):
                        row.cells[-1].text = f"教法：{methods.get('teacher_method', '情境引入、启发引导')}\n学法：{methods.get('student_method', '小组合作、自主探究')}"
                    else:
                        row.cells[-1].text = str(methods) if str(methods) else "教法：情境引入、启发引导\n学法：小组合作、自主探究"

                elif "归纳总结" in row_raw_text or "作业设计" in row_raw_text:
                    summary = word_data.get("summary_and_homework", {})
                    if isinstance(summary, dict):
                        row.cells[-1].text = f"【归纳总结】：{summary.get('summary', '')}\n【作业设计】：{summary.get('homework', '')}\n【素养发展】：{summary.get('literacy_development', '')}"
                    else:
                        row.cells[-1].text = str(summary) if str(summary) else f"归纳总结：梳理算理与算法。\n作业设计：完成课后练习并寻找生活中的小数。\n素养发展：提升数学运算与应用意识。"

                elif "板书设计" in row_raw_text:
                    # 避免写入到包含 ※板书设计※ 标签标题行，定位到下一行
                    if idx + 1 < len(table.rows):
                        table.rows[idx + 1].cells[-1].text = str(word_data.get("board_design", f"{lesson_title}\n1. 核心概念\n2. 计算步骤"))

                elif "课后反思" in row_raw_text:
                    # 定位到 ※课后反思※ 下一行
                    if idx + 1 < len(table.rows):
                        table.rows[idx + 1].cells[-1].text = str(word_data.get("reflection", "课堂结合本土素材，有效提升了学生的参与度与知识应用能力。"))

                # -------------------------------------------------------------
                # 3. 教学流程（五环节）精准匹配填充
                # -------------------------------------------------------------
                else:
                    first_cell_txt = row.cells[0].text.strip()
                    process_list = word_data.get('teaching_process', [])
                    
                    # 遍历生成的教学环节，与模板预设的行匹配
                    for p_item in process_list:
                        p_stage = p_item.get("stage", "")
                        # 判断当前行是否是模板的“创设情境”、“探究归纳建构新知”等行
                        if any(key in first_cell_txt and key in p_stage for key in ["创设", "探究", "理解", "迁移", "创新"]):
                            # 确定列序号，根据模板 [环节(0), 教师活动(1), 学生活动(2), 设计意图(3)]
                            if len(row.cells) >= 4:
                                row.cells[1].text = str(p_item.get("teacher_activity", ""))
                                row.cells[2].text = str(p_item.get("student_activity", ""))
                                row.cells[3].text = str(p_item.get("design_intent", ""))
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
