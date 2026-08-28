from flask import Flask, request, jsonify

app = Flask(__name__)

@app.route('/api/generate', methods=['POST'])
def generate():
    data = request.get_json(force=True) or {}

    lesson_title = data.get('lesson_title', '未命名课程')
    content = data.get('content', '')

    return jsonify({
        "status": "success",
        "docx_url": f"https://your-storage-domain.com/{lesson_title}.docx",
        "pptx_url": f"https://your-storage-domain.com/{lesson_title}.pptx"
    })

app = app
