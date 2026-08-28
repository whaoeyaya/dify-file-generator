from http.server import HTTPServer, BaseHTTPRequestHandler
import json
import os
import urllib.parse
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from pptx import Presentation
from pptx.util import Inches as PptInches, Pt as PptPt
from pptx.dml.color import RGBColor as PptRGBColor

DOWNLOAD_DIR = os.path.join(os.getcwd(), 'static')
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

def set_font(run, font_name='微软雅黑', size_pt=10.5, bold=False, color_rgb=(51, 51, 51)):
    run.font.name = font_name
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    rPr.append(rFonts)
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.font.color.rgb = RGBColor(*color_rgb)

def set_cell_bg(cell, hex_color="F7F9FA"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

class SimpleHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length) if content_length > 0 else b'{}'
            
            try:
                data = json.loads(post_data.decode('utf-8'))
            except Exception:
                data = {}

            summary = data.get('resource_summary', {})
            teaching_design = data.get('teaching_design', {})

            lesson_title = summary.get('lesson_title') or data.get('lesson_title') or '复式条形统计图'
            subject = summary.get('subject') or data.get('subject') or '数学'
            grade = summary.get('grade') or data.get('grade') or '五年级'
            teaching_hours = summary.get('teaching_hours') or data.get('teaching_hours') or '1课时'
            textbook_version = summary.get('textbook_version') or data.get('textbook_version') or '人教版'

            objectives = teaching_design.get('teaching_objectives') or data.get('teaching_objectives') or [
                '认识复式条形统计图，理解其实际应用价值。',
                '掌握复式条形统计图的绘制方法与步骤，能正确解读图表数据。',
                '经历数据收集、整理与分析全过程，提高数据分析观念。'
            ]

            process_list = teaching_design.get('teaching_process') or data.get('teaching_process') or [
                {"step_name": "创设情境", "teacher_activity": "出示单式条形统计图，提问如何更直观对比两组数据？", "student_activity": "观察思考，提出合并图表的想法。", "intent": "激发学习兴趣与对比需求。"},
                {"step_name": "探究归纳建构新知", "teacher_activity": "展示复式条形统计图，引导观察图例作用及直条绘制规则。", "student_activity": "讨论图例必要性，尝试绘制并总结规则。", "intent": "建立复式条形统计图概念。"},
                {"step_name": "知识的理解应用", "teacher_activity": "指导完成基础例题绘制并回答分析问题。", "student_activity": "独立完成统计图补充与数据解答。", "intent": "巩固图表绘制与读取能力。"},
                {"step_name": "知识的迁移应用", "teacher_activity": "提供本土/生活实际数据（如班级图书借阅情况）。", "student_activity": "分组分析趋势并做出推断。", "intent": "升华数据分析观念与应用能力。"},
                {"step_name": "知识的创新", "teacher_activity": "引导思考复式折线与复式条形图的异同与选用场景。", "student_activity": "展开讨论，发散思维。", "intent": "拓展知识结构与综合素养。"},
                {"step_name": "归纳总结 作业设计 素养发展", "teacher_activity": "引导课堂总结，布置基础与实践分层作业。", "student_activity": "总结收获，记录课后作业。", "intent": "内化知识，延伸课外实践。"}
            ]

            blackboard = teaching_design.get('blackboard_design') or data.get('blackboard_design') or "主板书：复式条形统计图（含标题、图例、直条标注）"
            reflection = teaching_design.get('reflection') or "本节课学生对图例的认知清晰，绘制环节需进一步加强工具使用的指导。"

            # -------------------------------------------------------------
            # 1. 严格按照图片模板渲染 Word 文档 (.docx)
            # -------------------------------------------------------------
            doc = Document()
            for s in doc.sections:
                s.top_margin = Inches(0.8)
                s.bottom_margin = Inches(0.8)
                s.left_margin = Inches(0.8)
                s.right_margin = Inches(0.8)

            # 顶部大标题
            p_head = doc.add_paragraph()
            p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_head = p_head.add_run(f"《{lesson_title}》教学设计")
            set_font(r_head, size_pt=18, bold=True, color_rgb=(31, 78, 121))

            # 创建 6 列的基础大表格
            table = doc.add_table(rows=0, cols=6)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.style = 'Table Grid'

            # 辅助函数：添加单行/合并行
            def add_row_info(k1, v1, k2, v2, k3, v3):
                row = table.add_row()
                cells = row.cells
                cells[0].text, cells[1].text = k1, v1
                cells[2].text, cells[3].text = k2, v2
                cells[4].text, cells[5].text = k3, v3
                for i in [0, 2, 4]:
                    set_cell_bg(cells[i], "F2F4F7")
                    set_font(cells[i].paragraphs[0].runs[0], bold=True)

            def add_merged_row(label, value=""):
                row = table.add_row()
                cells = row.cells
                # 合并 1~5 列
                a = cells[1].merge(cells[2]).merge(cells[3]).merge(cells[4]).merge(cells[5])
                cells[0].text = label
                set_cell_bg(cells[0], "F2F4F7")
                set_font(cells[0].paragraphs[0].runs[0], bold=True)
                if value:
                    a.text = value
                return a

            def add_section_header(title_text):
                row = table.add_row()
                cells = row.cells
                merged = cells[0].merge(cells[1]).merge(cells[2]).merge(cells[3]).merge(cells[4]).merge(cells[5])
                set_cell_bg(merged, "E8EEF5")
                p = merged.paragraphs[0]
                r = p.add_run(f"※{title_text}※")
                set_font(r, size_pt=11, bold=True, color_rgb=(31, 78, 121))

            # --- 填入基本信息组 ---
            add_row_info("授课教师", "", "授课教师单位", "", "授课日期", "")
            add_row_info("学段", "小学", "学科", subject, "适用年级", grade)
            add_row_info("授课时间", "40分钟", "课型", "新授课", "授课时数", teaching_hours)

            add_merged_row("课题", lesson_title)
            add_merged_row("教材分析", f"本课属于{textbook_version}教材，重点在于让学生掌握表达多组数据的方法。")
            add_merged_row("学情分析", "学生已掌握单式条形统计图的绘制，具备基础的数据解读能力。")
            
            obj_str = "\n".join([f"{i+1}. {o}" for i, o in enumerate(objectives)])
            add_merged_row("教学目标", obj_str)
            add_merged_row("教学重难点", "教学重点：理解复式条形统计图特点与绘制\n教学难点：图例的规范使用与数据分析")
            add_merged_row("教法与学法", "教法：启发式教学、情境教学法   学法：自主探究、合作交流")
            add_merged_row("教学流程", "创设情境 -> 探究新知 -> 理解应用 -> 迁移应用 -> 知识创新 -> 总结提升")

            # --- 填入【教学过程】模块（表头与多行内容） ---
            add_section_header("教学过程")

            # 教学过程表头行
            row_phead = table.add_row()
            cp = row_phead.cells
            cp[0].text = "教学环节"
            cp[1].text = "教师活动"
            cp[2].text = "学生活动"
            # cp[3], cp[4], cp[5] 合并为设计意图
            cp_intent = cp[3].merge(cp[4]).merge(cp[5])
            cp_intent.text = "设计意图"

            for i in [0, 1, 2]:
                set_cell_bg(cp[i], "EAEEF3")
                set_font(cp[i].paragraphs[0].runs[0], bold=True)
            set_cell_bg(cp_intent, "EAEEF3")
            set_font(cp_intent.paragraphs[0].runs[0], bold=True)

            # 填充教学环节各行
            for step in process_list:
                r_step = table.add_row()
                cs = r_step.cells
                cs[0].text = step.get('step_name', '')
                cs[1].text = step.get('teacher_activity', '')
                cs[2].text = step.get('student_activity', '')
                
                c_intent = cs[3].merge(cs[4]).merge(cs[5])
                c_intent.text = step.get('intent', '')

                set_cell_bg(cs[0], "F9FAFC")
                set_font(cs[0].paragraphs[0].runs[0], bold=True)

            # --- 填入【板书设计】与【课后反思】 ---
            add_section_header("板书设计")
            add_merged_row("板书设计", str(blackboard))

            add_section_header("课后反思")
            add_merged_row("课后反思", str(reflection))

            # 调整单元格内边距
            for row in table.rows:
                for cell in row.cells:
                    set_cell_margins(cell, top=120, bottom=120, left=150, right=150)

            docx_filename = f"{lesson_title}_教学设计表.docx"
            docx_path = os.path.join(DOWNLOAD_DIR, docx_filename)
            doc.save(docx_path)

            # -------------------------------------------------------------
            # 2. 生成多页 PPT 演示文稿 (.pptx)
            # -------------------------------------------------------------
            prs = Presentation()

            slide1 = prs.slides.add_slide(prs.slide_layouts[0])
            slide1.shapes.title.text = lesson_title
            slide1.placeholders[1].text = f"学科：{subject} | 年级：{grade} | 版本：{textbook_version}"

            slide2 = prs.slides.add_slide(prs.slide_layouts[1])
            slide2.shapes.title.text = "一、教学目标"
            tf2 = slide2.shapes.placeholders[1].text_frame
            tf2.text = ""
            for idx, obj in enumerate(objectives):
                p = tf2.add_paragraph() if idx > 0 else tf2.paragraphs[0]
                p.text = f"• {obj}"
                p.font.size = PptPt(18)

            for step in process_list:
                slide = prs.slides.add_slide(prs.slide_layouts[1])
                slide.shapes.title.text = f"环节：{step.get('step_name')}"
                tf = slide.shapes.placeholders[1].text_frame
                tf.text = f"【教师活动】\n{step.get('teacher_activity')}\n\n【学生活动】\n{step.get('student_activity')}"

            pptx_filename = f"{lesson_title}_课件.pptx"
            pptx_path = os.path.join(DOWNLOAD_DIR, pptx_filename)
            prs.save(pptx_path)

            # -------------------------------------------------------------
            # 3. 输出 JSON 格式响应
            # -------------------------------------------------------------
            host = self.headers.get('Host', '')
            protocol = 'https' if 'onrender.com' in host else 'http'
            base_url = f"{protocol}://{host}"

            response_data = {
                "status": "success",
                "message": "已按知识库表格标准生成 Word 与 PPT 文件！",
                "docx_url": f"{base_url}/download/{urllib.parse.quote(docx_filename)}",
                "pptx_url": f"{base_url}/download/{urllib.parse.quote(pptx_filename)}"
            }

            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps(response_data, ensure_ascii=False).encode('utf-8'))

        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({"error": str(e)}).encode('utf-8'))

    def do_GET(self):
        if self.path.startswith('/download/'):
            filename = urllib.parse.unquote(self.path.replace('/download/', ''))
            file_path = os.path.join(DOWNLOAD_DIR, filename)

            if os.path.exists(file_path):
                mime_type = 'application/vnd.openxmlformats-officedocument.presentationml.presentation' if filename.endswith('.pptx') else 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                self.send_response(200)
                self.send_header('Content-Type', mime_type)
                self.send_header('Content-Disposition', f"attachment; filename*=UTF-8''{urllib.parse.quote(filename)}")
                self.end_headers()
                with open(file_path, 'rb') as f:
                    self.wfile.write(f.read())
                return
            else:
                self.send_response(404)
                self.end_headers()
                self.wfile.write(b"File not found")
                return

        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()
        self.wfile.write(json.dumps({"status": "Server Online!"}).encode('utf-8'))

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    server = HTTPServer(('0.0.0.0', port), SimpleHandler)
    server.serve_forever()
