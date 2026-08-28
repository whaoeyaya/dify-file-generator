from http.server import HTTPServer, BaseHTTPRequestHandler
import json
import os
import urllib.parse
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from pptx import Presentation
from pptx.util import Inches as PptInches, Pt as PptPt
from pptx.dml.color import RGBColor as PptRGBColor
from pptx.enum.text import PP_ALIGN

# 创建静态下载文件目录
DOWNLOAD_DIR = os.path.join(os.getcwd(), 'static')
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# Word 设置字体辅助函数
def set_font(run, font_name='微软雅黑', size_pt=11, bold=False, color_rgb=(51, 51, 51)):
    run.font.name = font_name
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    rPr.append(rFonts)
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.font.color.rgb = RGBColor(*color_rgb)

# Word 设置单元格背景色
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

# PPT 页面标题设置辅助函数
def add_ppt_title(slide, text):
    title_shape = slide.shapes.title
    title_shape.text = text
    p = title_shape.text_frame.paragraphs[0]
    p.font.name = '微软雅黑'
    p.font.size = PptPt(28)
    p.font.bold = True
    p.font.color.rgb = PptRGBColor(31, 78, 121)

class SimpleHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length) if content_length > 0 else b'{}'
            
            try:
                data = json.loads(post_data.decode('utf-8'))
            except Exception:
                data = {}

            # 兼容嵌套格式及直接输入格式
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
                {
                    "step_name": "情境导入",
                    "teacher_activity": "出示单式条形统计图（如男、女生体育成绩统计），提问如何在一张图中更直观对比两组数据？",
                    "student_activity": "观察单式统计图，尝试提出合并图表、统一刻度与增加图例的构想。"
                },
                {
                    "step_name": "探究新知",
                    "teacher_activity": "展示复式条形统计图，引导学生观察“图例”的作用，示范并讲解并列直条的绘制规则。",
                    "student_activity": "讨论图例的作用，动手在绘制纸上补充直条与标注，总结绘制复式统计图的三大步骤。"
                },
                {
                    "step_name": "巩固应用",
                    "teacher_activity": "提供本土化/生活化数据（如学校各年级图书借阅情况），引导学生绘制并分析变化趋势。",
                    "student_activity": "独立或合作绘制复式条形统计图，回答相关推断问题并分享结论。"
                },
                {
                    "step_name": "总结拓展",
                    "teacher_activity": "引导学生总结复式条形统计图与单式条形统计图的异同，布置分层作业。",
                    "student_activity": "回顾本节课收获，谈谈复式条形统计图在生活中的应用场景。"
                }
            ]

            homework_list = teaching_design.get('homework') or data.get('homework') or [
                '基础题：完成教材配套练习册“复式条形统计图”第1-2题。',
                '拓展题：调查家里近3个月的水费与电费支出，绘制一张复式条形统计图。'
            ]

            blackboard = teaching_design.get('blackboard_design') or data.get('blackboard_design') or (
                "复式条形统计图\n"
                "1. 特点：能同时反映两组或多组数据及对比情况\n"
                "2. 要素：标题、横轴、纵轴、直条、图例（关键！）\n"
                "3. 步骤：画轴标数 -> 绘制直条（区分颜色/花纹） -> 标注数据"
            )

            # -------------------------------------------------------------
            # 1. 生成 Word 教学设计方案 (.docx)
            # -------------------------------------------------------------
            doc = Document()
            for section in doc.sections:
                section.top_margin = Inches(1)
                section.bottom_margin = Inches(1)
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)

            p_title = doc.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_title = p_title.add_run(f"《{lesson_title}》教学设计")
            set_font(run_title, font_name='微软雅黑', size_pt=20, bold=True, color_rgb=(31, 78, 121))
            doc.add_paragraph()

            # 基本信息表
            table = doc.add_table(rows=2, cols=4)
            table.autofit = False
            headers = [("学科", subject), ("适用年级", grade), ("课时", teaching_hours), ("教材版本", textbook_version)]
            for idx, (k, v) in enumerate(headers):
                r = idx // 2
                c = (idx % 2) * 2
                cell_k = table.cell(r, c)
                set_cell_background(cell_k, "F2F2F2")
                pk = cell_k.paragraphs[0]
                run_k = pk.add_run(k)
                set_font(run_k, '微软雅黑', 10.5, bold=True, color_rgb=(51, 51, 51))
                cell_v = table.cell(r, c+1)
                pv = cell_v.paragraphs[0]
                run_v = pv.add_run(v)
                set_font(run_v, '微软雅黑', 10.5, bold=False, color_rgb=(80, 80, 80))

            doc.add_paragraph()

            # 教学目标
            p_h1 = doc.add_paragraph()
            run_h1 = p_h1.add_run("一、 教学目标")
            set_font(run_h1, '微软雅黑', 14, bold=True, color_rgb=(31, 78, 121))
            for obj in objectives:
                p_item = doc.add_paragraph()
                run_item = p_item.add_run(f"•  {obj}")
                set_font(run_item, '微软雅黑', 11, color_rgb=(60, 60, 60))

            doc.add_paragraph()

            # 教学过程
            p_h2 = doc.add_paragraph()
            run_h2 = p_h2.add_run("二、 教学过程")
            set_font(run_h2, '微软雅黑', 14, bold=True, color_rgb=(31, 78, 121))
            for idx, step in enumerate(process_list, 1):
                p_step = doc.add_paragraph()
                run_sname = p_step.add_run(f"环节{idx}：{step.get('step_name', '教学环节')}")
                set_font(run_sname, '微软雅黑', 12, bold=True, color_rgb=(41, 128, 185))
                
                p_act1 = doc.add_paragraph()
                run_a1 = p_act1.add_run(f"【教师活动】 {step.get('teacher_activity', '')}")
                set_font(run_a1, '微软雅黑', 10.5, color_rgb=(51, 51, 51))
                
                p_act2 = doc.add_paragraph()
                run_a2 = p_act2.add_run(f"【学生活动】 {step.get('student_activity', '')}")
                set_font(run_a2, '微软雅黑', 10.5, color_rgb=(100, 100, 100))

            doc.add_paragraph()

            # 板书设计
            p_h3 = doc.add_paragraph()
            run_h3 = p_h3.add_run("三、 板书设计")
            set_font(run_h3, '微软雅黑', 14, bold=True, color_rgb=(31, 78, 121))
            p_bb = doc.add_paragraph()
            run_bb = p_bb.add_run(str(blackboard))
            set_font(run_bb, '微软雅黑', 11, color_rgb=(60, 60, 60))

            docx_filename = f"{lesson_title}_教学设计.docx"
            docx_path = os.path.join(DOWNLOAD_DIR, docx_filename)
            doc.save(docx_path)

            # -------------------------------------------------------------
            # 2. 生成多页完整的 PPT 演示文稿 (.pptx)
            # -------------------------------------------------------------
            prs = Presentation()

            # Slide 1: 封面页
            slide_layout = prs.slide_layouts[0]
            slide1 = prs.slides.add_slide(slide_layout)
            slide1.shapes.title.text = lesson_title
            slide1.placeholders[1].text = f"学科：{subject}   |   年级：{grade}   |   版本：{textbook_version}"

            # Slide 2: 教学目标页
            slide2 = prs.slides.add_slide(prs.slide_layouts[1])
            add_ppt_title(slide2, "一、教学目标")
            tf2 = slide2.shapes.placeholders[1].text_frame
            tf2.text = ""
            for idx, obj in enumerate(objectives):
                p = tf2.add_paragraph() if idx > 0 else tf2.paragraphs[0]
                p.text = f"•  {obj}"
                p.font.name = '微软雅黑'
                p.font.size = PptPt(18)
                p.font.color.rgb = PptRGBColor(60, 60, 60)
                p.space_after = PptPt(14)

            # Slide 3 ~ N: 教学过程各环节（每环节生成独立一页 PPT）
            for step in process_list:
                slide = prs.slides.add_slide(prs.slide_layouts[1])
                add_ppt_title(slide, f"教学环节：{step.get('step_name', '探究流程')}")
                tf = slide.shapes.placeholders[1].text_frame
                tf.text = ""

                p1 = tf.paragraphs[0]
                p1.text = "【教师活动】"
                p1.font.name = '微软雅黑'
                p1.font.size = PptPt(18)
                p1.font.bold = True
                p1.font.color.rgb = PptRGBColor(41, 128, 185)

                p1_sub = tf.add_paragraph()
                p1_sub.text = f"{step.get('teacher_activity', '')}"
                p1_sub.font.name = '微软雅黑'
                p1_sub.font.size = PptPt(16)
                p1_sub.font.color.rgb = PptRGBColor(60, 60, 60)
                p1_sub.space_after = PptPt(18)

                p2 = tf.add_paragraph()
                p2.text = "【学生活动】"
                p2.font.name = '微软雅黑'
                p2.font.size = PptPt(18)
                p2.font.bold = True
                p2.font.color.rgb = PptRGBColor(39, 174, 96)

                p2_sub = tf.add_paragraph()
                p2_sub.text = f"{step.get('student_activity', '')}"
                p2_sub.font.name = '微软雅黑'
                p2_sub.font.size = PptPt(16)
                p2_sub.font.color.rgb = PptRGBColor(60, 60, 60)

            # Slide N+1: 板书设计页
            slide_bb = prs.slides.add_slide(prs.slide_layouts[1])
            add_ppt_title(slide_bb, "板书设计")
            tf_bb = slide_bb.shapes.placeholders[1].text_frame
            tf_bb.text = ""
            for idx, line in enumerate(str(blackboard).split('\n')):
                if not line.strip():
                    continue
                p = tf_bb.add_paragraph() if idx > 0 else tf_bb.paragraphs[0]
                p.text = line.strip()
                p.font.name = '微软雅黑'
                p.font.size = PptPt(18)
                p.font.color.rgb = PptRGBColor(60, 60, 60)
                p.space_after = PptPt(10)

            # Slide N+2: 作业布置页
            slide_hw = prs.slides.add_slide(prs.slide_layouts[1])
            add_ppt_title(slide_hw, "课后作业")
            tf_hw = slide_hw.shapes.placeholders[1].text_frame
            tf_hw.text = ""
            for idx, hw in enumerate(homework_list):
                p = tf_hw.add_paragraph() if idx > 0 else tf_hw.paragraphs[0]
                p.text = f"•  {hw}"
                p.font.name = '微软雅黑'
                p.font.size = PptPt(18)
                p.font.color.rgb = PptRGBColor(60, 60, 60)
                p.space_after = PptPt(14)

            pptx_filename = f"{lesson_title}_课件.pptx"
            pptx_path = os.path.join(DOWNLOAD_DIR, pptx_filename)
            prs.save(pptx_path)

            # -------------------------------------------------------------
            # 3. 构造输出响应 URL
            # -------------------------------------------------------------
            host = self.headers.get('Host', '')
            protocol = 'https' if 'onrender.com' in host else 'http'
            base_url = f"{protocol}://{host}"

            docx_encoded = urllib.parse.quote(docx_filename)
            pptx_encoded = urllib.parse.quote(pptx_filename)

            response_data = {
                "status": "success",
                "message": "Word 与多页 PPT 课件已成功生成！",
                "docx_url": f"{base_url}/download/{docx_encoded}",
                "pptx_url": f"{base_url}/download/{pptx_encoded}"
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
                # 根据后缀设置 Content-Type
                if filename.endswith('.pptx'):
                    mime_type = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
                else:
                    mime_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

                self.send_response(200)
                self.send_header('Content-Type', mime_type)
                encoded_filename = urllib.parse.quote(filename)
                self.send_header('Content-Disposition', f"attachment; filename*=UTF-8''{encoded_filename}")
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
    print(f"Starting server on port {port}...")
    server.serve_forever()
