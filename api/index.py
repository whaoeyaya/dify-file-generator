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
from pptx.util import Pt as PptPt
from pptx.dml.color import RGBColor as PptRGBColor

DOWNLOAD_DIR = os.path.join(os.getcwd(), 'static')
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# -------------------------------------------------------------
# 自动定位 PPT 模板路径 (优先检查 api/ 目录，再检查根目录)
# -------------------------------------------------------------
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(CURRENT_DIR)

TEMPLATE_PATH = os.path.join(CURRENT_DIR, 'template.pptx')
if not os.path.exists(TEMPLATE_PATH):
    TEMPLATE_PATH = os.path.join(PARENT_DIR, 'template.pptx')

# -------------------------------------------------------------
# Word 修复字体函数：彻底解决 Word 中文显示为方框/乱码问题
# -------------------------------------------------------------
def apply_text_with_font(paragraph_or_cell, text, font_name='微软雅黑', size_pt=10.5, bold=False, color_rgb=(51, 51, 51)):
    p = paragraph_or_cell.paragraphs[0] if hasattr(paragraph_or_cell, 'paragraphs') else paragraph_or_cell
    p.text = "" 
    run = p.add_run(text)
    
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.font.color.rgb = RGBColor(*color_rgb)
    
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:cs'), font_name)
    rPr.append(rFonts)

def set_cell_bg(cell, hex_color="F7F9FA"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_ppt_title(slide, text):
    if slide.shapes.title:
        slide.shapes.title.text = text
        p = slide.shapes.title.text_frame.paragraphs[0]
        p.font.name = '微软雅黑'
        p.font.bold = True

class SimpleHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length) if content_length > 0 else b'{}'
            
            try:
                data = json.loads(post_data.decode('utf-8'))
            except Exception:
                data = {}

            # 兼容嵌套格式与平铺格式
            summary = data.get('resource_summary', {})
            teaching_design = data.get('teaching_design', {})

            lesson_title = summary.get('lesson_title') or data.get('lesson_title') or '复式条形统计图'
            subject = summary.get('subject') or data.get('subject') or '数学'
            grade = summary.get('grade') or data.get('grade') or '四年级'
            location = summary.get('location') or data.get('location') or '广西来宾市兴宾区'
            teaching_hours = summary.get('teaching_hours') or data.get('teaching_hours') or '1课时'
            textbook_version = summary.get('textbook_version') or data.get('textbook_version') or '人教版'

            objectives = teaching_design.get('teaching_objectives') or data.get('teaching_objectives') or [
                '知识与技能：结合来宾市双产业发展数据，掌握复式条形统计图的绘制方法。',
                '过程与方法：经历数据整理分析过程，体会统计在实际生活中的决策应用。',
                '本土素养：通过甘蔗与桑蚕产量分析，增强对家乡产业发展的自豪感与认同感。'
            ]

            process_list = teaching_design.get('teaching_process') or data.get('teaching_process') or [
                {"step_name": "创设情境", "teacher_activity": "展示来宾城厢镇与凤凰镇的甘蔗产量单式图，提问如何直观对比？", "student_activity": "观察思考，提出合并统计图的需求。", "intent": "触发本土认知冲突，引入复式概念。"},
                {"step_name": "探究新知", "teacher_activity": "演示两镇“甘蔗与蚕茧”双产业复式统计图，解析图例与直条规范。", "student_activity": "合作讨论图例作用，归纳绘制要点。", "intent": "建立复式条形统计图表征模型。"},
                {"step_name": "理解应用", "teacher_activity": "出示良江镇与小平阳镇农业数据，指导补全统计图。", "student_activity": "独立完成绘制，回答问题链。", "intent": "巩固图例绘制与基础解读能力。"},
                {"step_name": "迁移应用", "teacher_activity": "分析近3年来宾市蔗糖加工量的变化趋势。", "student_activity": "分组讨论并预测下季产量，撰写简短分析报告。", "intent": "提升高阶数据分析与推理素养。"},
                {"step_name": "创新升华", "teacher_activity": "引导讨论复式条形图与复式折线图在农作物生长监测中的选用。", "student_activity": "辩论异同，提出本土智慧农业表达方案。", "intent": "激发跨学科创新思维。"}
            ]

            ppt_slides_data = data.get('ppt_slides') or [
                {"title": "本土情境导入", "points": "对比城厢镇与凤凰镇甘蔗产量", "image_tip": "来宾蔗海风光与甘蔗运输车图", "interactive": "提问：如何在一个图里对比两个镇？"},
                {"title": "新知探究：认识图例", "points": "复式条形统计图的两个核心：图例与双直条", "image_tip": "标注图例和不同颜色直条的统计图演示", "interactive": "同桌讨论：为什么必须有图例？"},
                {"title": "动手绘制与实践", "points": "根据良江镇蚕茧与甘蔗数据补充图表", "image_tip": "未完成的统计图模板卡", "interactive": "学生动手绘制，拍照展示点评"},
                {"title": "数据分析与产业预测", "points": "根据图形判断哪种农作物增长更快", "image_tip": "来宾糖业加工产业链示意图", "interactive": "小组讨论：为农户提出种植建议"}
            ]

            homework_list = teaching_design.get('homework') or data.get('homework') or [
                '基础作业：完成教材配套练习册对应习题。',
                '本土迁移作业：调查自家近半年的水费与电费支出，绘制复式条形统计图并提出节约建议。'
            ]

            blackboard = teaching_design.get('blackboard_design') or data.get('blackboard_design') or "主板书：复式条形统计图\n1. 标题\n2. 图例（区分不同类别）\n3. 横轴与纵轴\n4. 双直条绘制（注意间距）"
            reflection = teaching_design.get('reflection') or "结合本土真实农产品数据显著提升了学生的参与度，后半段图例绘制细节仍需巡视强化指导。"

            # -------------------------------------------------------------
            # 1. 生成 Word 教学设计表格 (.docx)
            # -------------------------------------------------------------
            doc = Document()
            for s in doc.sections:
                s.top_margin = Inches(0.8)
                s.bottom_margin = Inches(0.8)
                s.left_margin = Inches(0.8)
                s.right_margin = Inches(0.8)

            p_head = doc.add_paragraph()
            p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
            apply_text_with_font(p_head, f"《{lesson_title}》本土化教学设计", size_pt=18, bold=True, color_rgb=(31, 78, 121))

            table = doc.add_table(rows=0, cols=6)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.style = 'Table Grid'

            def add_row_info(k1, v1, k2, v2, k3, v3):
                row = table.add_row()
                cs = row.cells
                pairs = [(cs[0], k1, True), (cs[1], v1, False), (cs[2], k2, True), (cs[3], v2, False), (cs[4], k3, True), (cs[5], v3, False)]
                for cell, txt, is_key in pairs:
                    apply_text_with_font(cell, txt, bold=is_key)
                    if is_key:
                        set_cell_bg(cell, "F2F4F7")

            def add_merged_row(label, value=""):
                row = table.add_row()
                cs = row.cells
                merged_cell = cs[1].merge(cs[2]).merge(cs[3]).merge(cs[4]).merge(cs[5])
                apply_text_with_font(cs[0], label, bold=True)
                set_cell_bg(cs[0], "F2F4F7")
                apply_text_with_font(merged_cell, value, bold=False)

            def add_section_header(title_text):
                row = table.add_row()
                cs = row.cells
                merged = cs[0].merge(cs[1]).merge(cs[2]).merge(cs[3]).merge(cs[4]).merge(cs[5])
                set_cell_bg(merged, "E8EEF5")
                apply_text_with_font(merged, f"※ {title_text} ※", size_pt=11, bold=True, color_rgb=(31, 78, 121))

            add_row_info("授课教师", "", "所属县域/学校", location, "授课日期", "")
            add_row_info("学段", "小学", "学科", subject, "适用年级", grade)
            add_row_info("授课时间", "40分钟", "课型", "新授课", "授课时数", teaching_hours)

            add_merged_row("课题名称", lesson_title)
            add_merged_row("教材及版本", f"{textbook_version} 结合区域特色素材重构")
            
            obj_str = "\n".join([f"{o}" if o.startswith(('1','2','3','知识','过程','本土')) else f"• {o}" for o in objectives])
            add_merged_row("教学目标", obj_str)
            add_merged_row("教学重难点", "教学重点：理解复式条形统计图特点，掌握图例规范绘制。\n教学难点：结合真实数据进行合理趋势推断与决策分析。")

            add_section_header("五环节教学流程")
            row_phead = table.add_row()
            cp = row_phead.cells
            apply_text_with_font(cp[0], "教学环节", bold=True)
            apply_text_with_font(cp[1], "教师活动", bold=True)
            apply_text_with_font(cp[2], "学生活动", bold=True)
            cp_intent = cp[3].merge(cp[4]).merge(cp[5])
            apply_text_with_font(cp_intent, "设计意图", bold=True)

            for cell in [cp[0], cp[1], cp[2], cp_intent]:
                set_cell_bg(cell, "EAEEF3")

            for step in process_list:
                r_step = table.add_row()
                cs = r_step.cells
                apply_text_with_font(cs[0], step.get('step_name', ''), bold=True)
                apply_text_with_font(cs[1], step.get('teacher_activity', ''))
                apply_text_with_font(cs[2], step.get('student_activity', ''))
                
                c_intent = cs[3].merge(cs[4]).merge(cs[5])
                apply_text_with_font(c_intent, step.get('intent', ''))
                set_cell_bg(cs[0], "F9FAFC")

            add_section_header("板书设计")
            add_merged_row("板书设计", str(blackboard))

            add_section_header("课后作业与拓展")
            hw_str = "\n".join([f"• {hw}" for hw in homework_list])
            add_merged_row("分层作业", hw_str)

            add_section_header("教学反思")
            add_merged_row("课后反思", str(reflection))

            for row in table.rows:
                for cell in row.cells:
                    set_cell_margins(cell, top=120, bottom=120, left=150, right=150)

            docx_filename = f"{lesson_title}_本土教学设计表.docx"
            docx_path = os.path.join(DOWNLOAD_DIR, docx_filename)
            doc.save(docx_path)

            # -------------------------------------------------------------
            # 2. 基于模板生成多页完整 PPT 课件 (.pptx)
            # -------------------------------------------------------------
            if os.path.exists(TEMPLATE_PATH):
                prs = Presentation(TEMPLATE_PATH)
                rId_list = [slide.rId for slide in prs.slides._sldIdLst]
                for rId in rId_list:
                    prs.part.drop_rel(rId)
                    del prs.slides._sldIdLst[0]
            else:
                prs = Presentation()

            layouts_count = len(prs.slide_layouts)
            layout_title = prs.slide_layouts[0] if layouts_count > 0 else prs.slide_layouts[0]
            layout_content = prs.slide_layouts[1] if layouts_count > 1 else prs.slide_layouts[0]

            # Slide 1: 封面页
            slide1 = prs.slides.add_slide(layout_title)
            if slide1.shapes.title:
                slide1.shapes.title.text = lesson_title
            if len(slide1.placeholders) > 1:
                slide1.placeholders[1].text = f"学科：{subject}   |   年级：{grade}   |   区域：{location}"

            # Slide 2: 教学目标页
            slide2 = prs.slides.add_slide(layout_content)
            add_ppt_title(slide2, "一、教学目标")
            if len(slide2.placeholders) > 1:
                tf2 = slide2.placeholders[1].text_frame
                tf2.text = ""
                for idx, obj in enumerate(objectives):
                    p = tf2.add_paragraph() if idx > 0 else tf2.paragraphs[0]
                    p.text = f"•  {obj}"
                    p.font.name = '微软雅黑'
                    p.font.size = PptPt(18)
                    p.space_after = PptPt(12)

            # Slide 3~N: 动态生成的 PPT 核心文案页
            for page in ppt_slides_data:
                slide = prs.slides.add_slide(layout_content)
                add_ppt_title(slide, page.get('title', '课堂探究'))
                if len(slide.placeholders) > 1:
                    tf = slide.placeholders[1].text_frame
                    tf.text = ""

                    p1 = tf.paragraphs[0]
                    p1.text = "【核心要点】"
                    p1.font.name = '微软雅黑'
                    p1.font.size = PptPt(18)
                    p1.font.bold = True
                    p1.font.color.rgb = PptRGBColor(41, 128, 185)

                    p1_sub = tf.add_paragraph()
                    p1_sub.text = page.get('points', '')
                    p1_sub.font.name = '微软雅黑'
                    p1_sub.font.size = PptPt(16)
                    p1_sub.space_after = PptPt(14)

                    if 'image_tip' in page:
                        p2 = tf.add_paragraph()
                        p2.text = "【配图/媒体建议】"
                        p2.font.name = '微软雅黑'
                        p2.font.size = PptPt(18)
                        p2.font.bold = True
                        p2.font.color.rgb = PptRGBColor(39, 174, 96)

                        p2_sub = tf.add_paragraph()
                        p2_sub.text = page.get('image_tip', '')
                        p2_sub.font.name = '微软雅黑'
                        p2_sub.font.size = PptPt(16)
                        p2_sub.space_after = PptPt(14)

                    if 'interactive' in page:
                        p3 = tf.add_paragraph()
                        p3.text = "【互动提示】"
                        p3.font.name = '微软雅黑'
                        p3.font.size = PptPt(18)
                        p3.font.bold = True
                        p3.font.color.rgb = PptRGBColor(211, 84, 0)

                        p3_sub = tf.add_paragraph()
                        p3_sub.text = page.get('interactive', '')
                        p3_sub.font.name = '微软雅黑'
                        p3_sub.font.size = PptPt(16)

            # 结尾页: 课后作业页
            slide_hw = prs.slides.add_slide(layout_content)
            add_ppt_title(slide_hw, "课后作业与实践")
            if len(slide_hw.placeholders) > 1:
                tf_hw = slide_hw.placeholders[1].text_frame
                tf_hw.text = ""
                for idx, hw in enumerate(homework_list):
                    p = tf_hw.add_paragraph() if idx > 0 else tf_hw.paragraphs[0]
                    p.text = f"•  {hw}"
                    p.font.name = '微软雅黑'
                    p.font.size = PptPt(18)
                    p.space_after = PptPt(14)

            pptx_filename = f"{lesson_title}_课件.pptx"
            pptx_path = os.path.join(DOWNLOAD_DIR, pptx_filename)
            prs.save(pptx_path)

            # -------------------------------------------------------------
            # 3. 输出 Json 响应
            # -------------------------------------------------------------
            host = self.headers.get('Host', '')
            protocol = 'https' if 'onrender.com' in host else 'http'
            base_url = f"{protocol}://{host}"

            response_data = {
                "status": "success",
                "message": "Word 与多页模板 PPT 课件已成功生成！",
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
            self.send_header('Access-Control-Allow-Origin', '*')
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
                self.send_header('Access-Control-Allow-Origin', '*')
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
