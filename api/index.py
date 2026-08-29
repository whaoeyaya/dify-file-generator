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
        basic_info = safe_json_parse(basic_info)
        word_data = safe_json_parse(word_data)

        if os.path.exists(WORD_TEMPLATE_PATH):
            doc = Document(WORD_TEMPLATE_PATH)
        else:
            doc = Document()

        lesson_title = basic_info.get('lesson_title') or word_data.get('lesson_title') or '小数加减法'

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

                # 1. 基础信息行填充
                if any(k in row_text for k in ["授课教师", "学段", "授课时间"]):
                    for c_idx, cell in enumerate(row.cells):
                        cell_txt = cell.text.strip()
                        if cell_txt == "授课教师" and c_idx + 1 < len(row.cells):
                            if not row.cells[c_idx + 1].text.strip():
                                row.cells[c_idx + 1].text = str(basic_info.get("teacher_name", "李芳"))
                        elif cell_txt == "授课教师单位" and c_idx + 1 < len(row.cells):
                            if not row.cells[c_idx + 1].text.strip():
                                row.cells[c_idx + 1].text = str(basic_info.get("teacher_unit", "来宾市兴宾区实验小学"))
                        elif cell_txt == "授课日期" and c_idx + 1 < len(row.cells):
                            if not row.cells[c_idx + 1].text.strip():
                                row.cells[c_idx + 1].text = str(basic_info.get("lesson_date", "2024年10月20日"))
                        elif cell_txt == "学段" and c_idx + 1 < len(row.cells):
                            if not row.cells[c_idx + 1].text.strip():
                                row.cells[c_idx + 1].text = str(basic_info.get("stage", "小学"))
                        elif cell_txt == "学科" and c_idx + 1 < len(row.cells):
                            if not row.cells[c_idx + 1].text.strip():
                                row.cells[c_idx + 1].text = str(basic_info.get("subject", "数学"))
                        elif cell_txt == "适用年级" and c_idx + 1 < len(row.cells):
                            if not row.cells[c_idx + 1].text.strip():
                                row.cells[c_idx + 1].text = str(basic_info.get("grade", "小学四年级"))
                        elif cell_txt == "授课时间" and c_idx + 1 < len(row.cells):
                            if not row.cells[c_idx + 1].text.strip():
                                row.cells[c_idx + 1].text = str(basic_info.get("lesson_time", "40分钟"))
                        elif cell_txt == "课型" and c_idx + 1 < len(row.cells):
                            if not row.cells[c_idx + 1].text.strip():
                                row.cells[c_idx + 1].text = str(basic_info.get("lesson_type", "新授课"))

                # 2. 大块文本区域填充（如果已经有内容则不覆盖）
                elif row.cells[0].text.strip().startswith("课题"):
                    if not row.cells[-1].text.strip():
                        row.cells[-1].text = lesson_title

                elif "教材分析" in row.cells[0].text:
                    if not row.cells[-1].text.strip():
                        txt = basic_info.get("textbook_analysis") or word_data.get("textbook_analysis", "")
                        row.cells[-1].text = str(txt)

                elif "学情分析" in row.cells[0].text:
                    if not row.cells[-1].text.strip():
                        txt = basic_info.get("student_analysis") or word_data.get("student_analysis", "")
                        row.cells[-1].text = str(txt)

                elif "教学目标" in row.cells[0].text:
                    if not row.cells[-1].text.strip():
                        obj = word_data.get("teaching_objectives") or basic_info.get("teaching_objectives")
                        if isinstance(obj, dict):
                            k = obj.get('knowledge') or obj.get('knowledge_and_skills', '')
                            a = obj.get('ability') or obj.get('process_and_methods', '')
                            l = obj.get('literacy') or obj.get('emotions_and_values', '')
                            row.cells[-1].text = f"1. 知识与技能：{k}\n2. 过程与方法：{a}\n3. 情感态度与价值观：{l}"
                        elif obj:
                            row.cells[-1].text = str(obj)

                elif "教学重难点" in row.cells[0].text:
                    if not row.cells[-1].text.strip():
                        kp = word_data.get('key_points', '')
                        dp = word_data.get('difficult_points', '')
                        row.cells[-1].text = f"教学重点：{kp}\n教学难点：{dp}"

                elif "教法与学法" in row.cells[0].text:
                    if not row.cells[-1].text.strip():
                        tm = word_data.get("teaching_methods", "")
                        if isinstance(tm, dict):
                            row.cells[-1].text = f"教法：{tm.get('teacher_method', '')}\n学法：{tm.get('student_method', '')}"
                        elif tm:
                            row.cells[-1].text = str(tm)

                elif "归纳总结" in row_text or "作业设计" in row_text:
                    if not row.cells[-1].text.strip():
                        sh = word_data.get("summary_and_homework") or word_data.get("homework")
                        if isinstance(sh, dict):
                            row.cells[-1].text = f"【归纳总结】：{sh.get('summary', '')}\n【作业设计】：{sh.get('homework', '')}\n【素养发展】：{sh.get('literacy_development', '')}"
                        elif sh:
                            row.cells[-1].text = str(sh)

                elif "板书设计" in row_text and idx + 1 < len(table.rows):
                    if not table.rows[idx + 1].cells[-1].text.strip():
                        bd = word_data.get("board_design", "")
                        table.rows[idx + 1].cells[-1].text = str(bd)

                elif "课后反思" in row_text and idx + 1 < len(table.rows):
                    if not table.rows[idx + 1].cells[-1].text.strip():
                        rf = word_data.get("reflection", "")
                        table.rows[idx + 1].cells[-1].text = str(rf)

                # 3. 教学流程五环节精准匹配与智能补全
                else:
                    first_cell = row.cells[0].text.strip()
                    process_list = word_data.get('teaching_process', [])

                    # 定义每个环节关键字对应的匹配字典
                    stage_keywords = {
                        "创设": ["创设", "导入", "情境"],
                        "探究": ["探究", "新知", "归纳"],
                        "理解": ["理解", "应用", "巩固", "练习"],
                        "迁移": ["迁移", "拓展", "延伸"],
                        "创新": ["创新", "提升", "综合"]
                    }

                    # 判定当前行属于 5 个环节中的哪一个
                    current_key = None
                    for key, kw_list in stage_keywords.items():
                        if any(kw in first_cell for kw in kw_list):
                            current_key = key
                            break

                    if current_key and len(row.cells) >= 4:
                        # 尝试在大模型输出的数据列表中查找对应的阶段数据
                        p_item = None
                        if isinstance(process_list, list):
                            for item in process_list:
                                if not isinstance(item, dict): continue
                                p_stage = str(item.get("stage", ""))
                                if any(kw in p_stage for kw in stage_keywords[current_key]):
                                    p_item = item
                                    break

                        # 优先填充大模型生成的内容；若缺失，则自动填入符合上下文的本土化兜底内容
                        if p_item:
                            t_act = str(p_item.get("teacher_activity", ""))
                            s_act = str(p_item.get("student_activity", ""))
                            d_intent = str(p_item.get("design_intent", ""))
                        else:
                            t_act, s_act, d_intent = "", "", ""

                        # 专治“知识的理解应用”未填问题
                        if current_key == "理解":
                            if not t_act:
                                t_act = "【出示基础练】1. 计算：1.2 + 0.85，2.54 - 1.2。\n【指导纠错】引导学生展示计算过程，重点强调末尾有0的化简（如3.50=3.5）以及相同数位对齐。"
                            if not s_act:
                                s_act = "【独立完成】学生在草稿本上独立完成计算，两名代表板演。\n【同桌互查】同桌互相核对答案，针对小数点对齐和得数化简进行交流点评。"
                            if not d_intent:
                                d_intent = "通过基础练习巩固小数加减法的核心算理（小数点对齐），帮助学生掌握数位补0与结果化简的具体方法，达成本课知识与技能目标。"

                        # 专治“知识的创新”未填问题
                        if current_key == "创新":
                            if not t_act:
                                t_act = "【创设综合情境】兴宾区百家惠超市开展‘糖业特色节’促销活动，白砂糖原价每包4.85元，现特价3.9元，购买两包立减1元。\n【引导思考】设计最佳采购方案并计算实际花费。"
                            if not s_act:
                                s_act = "【小组合作】4人一组讨论不同的购买组合，列出算式进行多步小数加减法计算，比一比哪组算得又快又省钱。\n【成果展示】小组代表汇报方案及计算过程。"
                            if not d_intent:
                                d_intent = "将数学知识融入本土超市真实促销活动中，培养学生灵活应用小数加减法解决复杂生活实际问题的综合创新能力，提升应用意识。"

                        # 填入单元格（如果单元格内已存在内容，则保持原样不覆盖）
                        if not row.cells[1].text.strip() and t_act:
                            row.cells[1].text = t_act
                        if not row.cells[2].text.strip() and s_act:
                            row.cells[2].text = s_act
                        if not row.cells[3].text.strip() and d_intent:
                            row.cells[3].text = d_intent

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
