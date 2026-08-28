from http.server import BaseHTTPRequestHandler
import json

class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            # 读取请求体长度并获取数据
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length) if content_length > 0 else b'{}'
            
            # 解析安全 JSON
            try:
                data = json.loads(post_data.decode('utf-8'))
            except Exception:
                data = {}

            lesson_title = data.get('lesson_title', '教学设计')

            # 模拟返回的 JSON 结构（填入后续你要返回的下载链接逻辑）
            response_data = {
                "status": "success",
                "docx_url": f"https://your-domain.com/downloads/{lesson_title}.docx",
                "pptx_url": f"https://your-domain.com/downloads/{lesson_title}.pptx"
            }

            # 正常返回 200 响应
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
        # GET 请求保底，方便浏览器直接测试
        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()
        self.wfile.write(json.dumps({"message": "Serverless API is Online!"}).encode('utf-8'))

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
