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
    run = p.add_run(str(text) if text is not None else "")
    
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

# -------------------------------------------------------------
# 辅助提取函数：兼容大模型输出的多种 JSON 嵌套格式
# -------------------------------------------------------------
def get_field_from_data(data, keys, default=""):
    """尝试从字典或嵌套字典中按优先级提取字段"""
    for key_path in keys:
        temp = data
        found = True
        for k in key_path.split('.'):
            if isinstance(temp, dict) and k in temp:
                temp = temp[k]
            else:
                found = False
                break
        if found and temp:
            return temp
    return default

class SimpleHandler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        """处理跨域预检请求"""
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_POST(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length) if content_length > 0 else b'{}'
            
            try:
                data = json.loads(post_data.decode('utf-8'))
            except Exception:
                data = {}

            # 如果数据被嵌套在 body / payload 中，自动解包
            if 'body' in data and isinstance(data['body'], dict):
                data = data['body']

            summary = data.get('resource_summary', {})
            teaching_design = data.get('teaching_design', {})
            basic_info = teaching_design.get('basic_info', {}) if isinstance(teaching_design, dict) else {}

            # 多路径智能提取通用变量
            lesson_title = get_field_from_data(data, [
                'resource_summary.title', 'resource_summary.lesson_title', 
                'teaching_design.basic_info.title', 'lesson_title', 'title'
            ], default='未命名课程设计')

            subject = get_field_from_data(data, [
                'resource_summary.subject', 'teaching_design.basic_info.subject', 'subject'
            ], default='通用学科')

            grade = get_field_from_data(data, [
                'resource_summary.grade', 'teaching_design.basic_info.grade', 'grade'
            ], default='通用年级')

            location = get_field_from_data(data, [
                'resource_summary.location', 'teaching_design.basic_info.location', 'location'
            ], default='当地学校')

            teaching_hours = get_field_from_data(data, [
                'resource_summary.teaching_hours', 'teaching_design.basic_info.teaching_hours', 
                'teaching_hours', 'time_allocation'
            ], default='1课时')

            textbook_version = get_field_from_data(data, [
                'resource_summary.textbook_version', 'teaching_design.basic_info.textbook_version', 'textbook_version'
            ], default='标准教材')

            # 提取教学目标 (支持列表或字符串)
            raw_objectives = get_field_from_data(data, [
                'teaching_design.objectives', 'teaching_design.teaching_objectives', 'teaching_objectives', 'objectives'
            ], default=[])
            
            if isinstance(raw_objectives, str):
                objectives = [o.strip() for o in raw_objectives.split('\n') if o.strip()]
            elif isinstance(raw_objectives, list):
                objectives = [str(o) for o in raw_objectives]
            else:
                objectives = ['核心素养与知识目标建构中...']

            # 提取教学过程 process_list
            raw_process = get_field_from_data(data, [
                'teaching_design.teaching_process', 'teaching_process', 'process_list'
            ], default=[])

            process_list = []
            if isinstance(raw_process, list):
                for item in raw_process:
                    if isinstance(item, dict):
                        process_list.append({
                            'step_name': item.get('stage_name') or item.get('step_name') or item.get('stage') or '教学环节',
                            'teacher_activity': item.get('teacher_activity') or item.get('teacher') or '',
                            'student_activity': item.get('student_activity') or item.get('student') or '',
                            'intent': item.get('design_intent') or item.get('intent') or item.get('local_material_integration') or ''
                        })

            if not process_list:
                process_list = [
                    {"step_name": "情境创设", "teacher_activity": "展示引入情境与素材", "student_activity": "观察并思考核心问题", "intent": "激发学习兴趣，建立课堂情境"},
                    {"step_name": "任务驱动", "teacher_activity": "布置探究任务与操作指引", "student_activity": "分组合作探究或独立思考", "intent": "探究新知，建构核心概念"},
                    {"step_name": "总结升华", "teacher_activity": "总结梳理方法与提炼知识点", "student_activity": "归纳交流，表达学习收获", "intent": "巩固知识体系，进行素养升华"}
                ]

            # 提取 PPT 页面数据
            raw_ppt_slides = get_field_from_data(data, [
                'ppt_slides', 'pptx_slides', 'ppt_structure'
            ], default=[])

            ppt_slides_data = []
            if isinstance(raw_ppt_slides, list):
                for slide in raw_ppt_slides:
                    if isinstance(slide, dict):
                        # 处理 main_content (支持数组或字符串)
                        main_c = slide.get('main_content') or slide.get('points') or ''
                        if isinstance(main_c, list):
                            main_c = "\n".join([f"• {x}" for x in main_c])
                        
                        ppt_slides_data.append({
                            'title': slide.get('slide_title') or slide.get('title') or '课堂探究',
                            'points': main_c,
                            'image_tip': slide.get('slide_purpose') or slide.get('image_tip') or '',
                            'interactive': slide.get('teacher_instruction') or slide.get('student_activity') or slide.get('interactive') or ''
                        })

            if not ppt_slides_data:
                ppt_slides_data = [
                    {"title": f"《{lesson_title}》情境导入", "points": "围绕主题开展教学设计与探究活动", "image_tip": "情境图片与教学媒体推荐", "interactive": "提问互动与思考"},
                    {"title": "核心知识点与探究", "points": "重点概念提炼与例题解析", "image_tip": "结构图与示意图展示", "interactive": "分组讨论与师生协同"}
                ]

            # 提取作业与板书
            homework_data = get_field_from_data(data, [
                'teaching_design.homework', 'homework', 'homework_list'
            ], default=[])
            if isinstance(homework_data, dict):
                homework_list = [f"{k}: {v}" for k, v in homework_data.items()]
            elif isinstance(homework_data, list):
                homework_list = [str(h) for h in homework_data]
            elif isinstance(homework_data, str):
                homework_list = [h.strip() for h in homework_data.split('\n') if h.strip()]
            else:
                homework_list = ['完成对应配套练习手册。']

            blackboard = get_field_from_data(data, [
                'teaching_design.board_design', 'teaching_design.blackboard_design', 'board_design', 'blackboard_design'
            ], default=f"板书设计：{lesson_title}\n1. 核心概念\n2. 探究方法\n3. 总结归纳")

            reflection = get_field_from_data(data, [
                'teaching_design.evaluation_and_reflection', 'teaching_design.reflection', 'reflection'
            ], default="根据课堂实际生成反馈，注重本土素材与核心概念的深度结合。")

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
            apply_text_with_font(p_head, f"《{lesson_title}》教学设计", size_pt=18, bold=True, color_rgb=(31, 78, 121))

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
            add_row_info("学段", "中小学", "学科", subject, "适用年级", grade)
            add_row_info("授课时间", "40分钟", "课型", "新授课", "授课时数", teaching_hours)

            add_merged_row("课题名称", lesson_title)
            add_merged_row("教材及版本", f"{textbook_version}")
            
            obj_str = "\n".join([f"• {o}" if not o.startswith('•') else o for o in objectives])
            add_merged_row("教学目标", obj_str)
            add_merged_row("教学重难点", "教学重点：掌握核心概念与方法。\n教学难点：灵活融会贯通并解决实际问题。")

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
            hw_str = "\n".join([f"• {hw}" if not hw.startswith('•') else hw for hw in homework_list])
            add_merged_row("分层作业", hw_str)

            add_section_header("教学反思")
            add_merged_row("课后反思", str(reflection))

            for row in table.rows:
                for cell in row.cells:
                    set_cell_margins(cell, top=120, bottom=120, left=150, right=150)

            # 安全文件名处理（去除路径不安全字符）
            safe_title = "".join([c for c in lesson_title if c.isalnum() or c in (' ', '_', '-')]).strip() or "教学设计"
            docx_filename = f"{safe_title}_教学设计表.docx"
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

                    if page.get('image_tip'):
                        p2 = tf.add_paragraph()
                        p2.text = "【设计意图/媒体建议】"
                        p2.font.name = '微软雅黑'
                        p2.font.size = PptPt(18)
                        p2.font.bold = True
                        p2.font.color.rgb = PptRGBColor(39, 174, 96)

                        p2_sub = tf.add_paragraph()
                        p2_sub.text = page.get('image_tip', '')
                        p2_sub.font.name = '微软雅黑'
                        p2_sub.font.size = PptPt(16)
                        p2_sub.space_after = PptPt(14)

                    if page.get('interactive'):
                        p3 = tf.add_paragraph()
                        p3.text = "【教学互动指引】"
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

            pptx_filename = f"{safe_title}_课件.pptx"
            pptx_path = os.path.join(DOWNLOAD_DIR, pptx_filename)
            prs.save(pptx_path)

            # -------------------------------------------------------------
            # 3. 输出 JSON 响应
            # -------------------------------------------------------------
            host = self.headers.get('Host', '')
            protocol = 'https' if 'onrender.com' in host else 'http'
            base_url = f"{protocol}://{host}"

            response_data = {
                "status": "success",
                "message": f"《{lesson_title}》Word 与 PPT 课件已成功动态生成！",
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
