from http.server import HTTPServer, BaseHTTPRequestHandler
import json
import os
import urllib.parse
from docx import Document
from pptx import Presentation

# 创建临时文件保存目录
DOWNLOAD_DIR = os.path.join(os.getcwd(), 'static')
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

class SimpleHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length) if content_length > 0 else b'{}'
            
            try:
                data = json.loads(post_data.decode('utf-8'))
            except Exception:
                data = {}

            lesson_title = data.get('lesson_title', '教学设计方案')
            content = data.get('content', '暂无详细教学内容')

            # 1. 生成 Word 文件
            doc = Document()
            doc.add_heading(lesson_title, 0)
            doc.add_paragraph(content)
            docx_filename = f"{lesson_title}.docx"
            docx_path = os.path.join(DOWNLOAD_DIR, docx_filename)
            doc.save(docx_path)

            # 2. 生成 PPT 文件
            prs = Presentation()
            slide_layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = lesson_title
            slide.placeholders[1].text = content[:100]  # 节选前100字作为概要
            pptx_filename = f"{lesson_title}.pptx"
            pptx_path = os.path.join(DOWNLOAD_DIR, pptx_filename)
            prs.save(pptx_path)

            # 3. 动态获取当前请求的完整域名，拼装真实下载链接
            host = self.headers.get('Host', '')
            protocol = 'https' if 'onrender.com' in host else 'http'
            base_url = f"{protocol}://{host}"

            docx_encoded = urllib.parse.quote(docx_filename)
            pptx_encoded = urllib.parse.quote(pptx_filename)

            response_data = {
                "status": "success",
                "message": "文件生成成功！",
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
        # 处理文件下载接口 /download/文件名
        if self.path.startswith('/download/'):
            filename = urllib.parse.unquote(self.path.replace('/download/', ''))
            file_path = os.path.join(DOWNLOAD_DIR, filename)

            if os.path.exists(file_path):
                self.send_response(200)
                self.send_header('Content-Type', 'application/octet-stream')
                self.send_header('Content-Disposition', f'attachment; filename="{urllib.parse.quote(filename)}"')
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
