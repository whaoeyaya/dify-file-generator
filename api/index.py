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

def safe_json_parse(data):
    """确保传入的数据转为 dict 格式"""
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
        # 1. 安全转码，防止传入的是 JSON 字符串
        basic_info = safe_json_parse(basic_info)
        word_data = safe_json_parse(word_data)

        # 2. 优先加载原始 template.docx 模板
        if os.path.exists(WORD_TEMPLATE_PATH):
            doc = Document(WORD_TEMPLATE_PATH)
        else:
            doc = Document()

        lesson_title = basic_info.get('lesson_title') or word_data.get('lesson_title') or '教学设计'

        # A. 替换文档主标题
        for p in doc.paragraphs:
            if "教学设计" in p.text:
                p.text = f"《{lesson_title}》教学设计"
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                break

        # B. 填充主表格
        if len(doc.tables) > 0:
            table = doc.tables[0]

            for idx, row in enumerate(table.rows):
                row_text = "".join([c.text.strip() for c in row.cells])

                # -------------------------------------------------------------
                # 1. 基础信息行填充
                # -------------------------------------------------------------
                if any(k in row_text for k in ["授课教师", "学段", "授课时间"]):
                    for c_idx, cell in enumerate(row.cells):
                        cell_txt = cell.text.strip()
                        if cell_txt == "授课教师" and c_idx + 1 < len(row.cells):
                            cell_val = basic_info.get("teacher_name") or "教师"
                            row.cells[c_idx + 1].text = str(cell_val)
                        elif cell_txt == "授课教师单位" and c_idx + 1 < len(row.cells):
                            cell_val = basic_info.get("teacher_unit") or "实验小学"
                            row.cells[c_idx + 1].text = str(cell_val)
                        elif cell_txt == "授课日期" and c_idx + 1 < len(row.cells):
                            cell_val = basic_info.get("lesson_date") or "2026年"
                            row.cells[c_idx + 1].text = str(cell_val)
                        elif cell_txt == "学段" and c_idx + 1 < len(row.cells):
                            cell_val = basic_info.get("stage") or "小学"
                            row.cells[c_idx + 1].text = str(cell_val)
                        elif cell_txt == "学科" and c_idx + 1 < len(row.cells):
                            cell_val = basic_info.get("subject") or "数学"
                            row.cells[c_idx + 1].text = str(cell_val)
                        elif cell_txt == "适用年级" and c_idx + 1 < len(row.cells):
                            cell_val = basic_info.get("grade") or "四年级"
                            row.cells[c_idx + 1].text = str(cell_val)
                        elif cell_txt == "授课时间" and c_idx + 1 < len(row.cells):
                            cell_val = basic_info.get("lesson_time") or "40分钟"
                            row.cells[c_idx + 1].text = str(cell_val)
                        elif cell_txt == "课型" and c_idx + 1 < len(row.cells):
                            cell_val = basic_info.get("lesson_type") or "新授课"
                            row.cells[c_idx + 1].text = str(cell_val)

                # -------------------------------------------------------------
                # 2. 大块文本区域填充（使用 row.cells[-1] 填充右侧合并单元格）
                # -------------------------------------------------------------
                elif row.cells[0].text.strip().startswith("课题"):
                    row.cells[-1].text = lesson_title

                elif "教材分析" in row.cells[0].text:
                    txt = basic_info.get("textbook_analysis") or word_data.get("textbook_analysis") or "结合本土实际素材深入浅出阐述概念。"
                    row.cells[-1].text = str(txt)

                elif "学情分析" in row.cells[0].text:
                    txt = basic_info.get("student_analysis") or word_data.get("student_analysis") or "学生具备基础计算能力，但需加强生活应用。"
                    row.cells[-1].text = str(txt)

                elif "教学目标" in row.cells[0].text:
                    obj = word_data.get("teaching_objectives") or basic_info.get("teaching_objectives")
                    if isinstance(obj, dict):
                        k = obj.get('knowledge') or obj.get('knowledge_and_skills') or '掌握基础计算方法。'
                        a = obj.get('ability') or obj.get('process_and_methods') or '通过小组合作提升问题解决能力。'
                        l = obj.get('literacy') or obj.get('emotions_and_values') or '体会数学在生活中的实用价值。'
                        row.cells[-1].text = f"1. 知识与技能：{k}\n2. 过程与方法：{a}\n3. 情感态度与价值观：{l}"
                    elif obj:
                        row.cells[-1].text = str(obj)

                elif "教学重难点" in row.cells[0].text:
                    kp = word_data.get('key_points') or '掌握计算法则及数位对齐。'
                    dp = word_data.get('difficult_points') or '理解跨位数计算及退位算理。'
                    row.cells[-1].text = f"教学重点：{kp}\n教学难点：{dp}"

                elif "教法与学法" in row.cells[0].text:
                    tm = word_data.get("teaching_methods")
                    if isinstance(tm, dict):
                        row.cells[-1].text = f"教法：{tm.get('teacher_method', '情境教学法')}\n学法：{tm.get('student_method', '自主探究、小组合作')}"
                    elif tm:
                        row.cells[-1].text = str(tm)
                    else:
                        row.cells[-1].text = "教法：情境教学法、启发引导\n学法：自主探究、合作交流"

                elif "归纳总结" in row_text or "作业设计" in row_text:
                    sh = word_data.get("summary_and_homework") or word_data.get("homework")
                    if isinstance(sh, dict):
                        row.cells[-1].text = f"【归纳总结】：{sh.get('summary', '梳理核心算理')}\n【作业设计】：{sh.get('homework', '完成课后练习')}\n【素养发展】：{sh.get('literacy_development', '提升运算能力')}"
                    elif sh:
                        row.cells[-1].text = str(sh)
                    else:
                        row.cells[-1].text = "【归纳总结】：总结相同数位对齐的计算方法。\n【作业设计】：结合本土超市购物结算完成计算。\n【素养发展】：培养认真严谨的计算习惯。"

                elif "板书设计" in row_text and idx + 1 < len(table.rows):
                    bd = word_data.get("board_design") or f"{lesson_title}\n1. 关键步骤\n2. 算理展示"
                    table.rows[idx + 1].cells[-1].text = str(bd)

                elif "课后反思" in row_text and idx + 1 < len(table.rows):
                    rf = word_data.get("reflection") or "通过本土化真实情境引入，学生参与积极性高，有效突破了重难点。"
                    table.rows[idx + 1].cells[-1].text = str(rf)

                # -------------------------------------------------------------
                # 3. 教学流程五环节填充
                # -------------------------------------------------------------
                else:
                    first_cell = row.cells[0].text.strip()
                    process_list = word_data.get('teaching_process', [])

                    if isinstance(process_list, list) and len(process_list) > 0:
                        for p_item in process_list:
                            if not isinstance(p_item, dict): continue
                            p_stage = str(p_item.get("stage", ""))

                            # 匹配模板左侧自带的 5 个流程阶段名称
                            matched = False
                            for kw in ["创设", "探究", "理解", "迁移", "创新"]:
                                if kw in first_cell and kw in p_stage:
                                    matched = True
                                    break

                            if matched and len(row.cells) >= 4:
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
