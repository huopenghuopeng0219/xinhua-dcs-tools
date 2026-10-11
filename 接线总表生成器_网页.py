# -*- coding: utf-8 -*-
"""
接线总表生成器 V1.0
网页界面版本 - 用浏览器操作
"""
import os
import sys
import webbrowser
import threading
from flask import Flask, render_template_string, request, jsonify

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from generate_wiring_table_new import WiringTableGenerator

app = Flask(__name__)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>接线总表生成器 V1.0</title>
    <style>
        body { font-family: 微软雅黑, sans-serif; max-width: 600px; margin: 50px auto; padding: 20px; }
        h1 { color: #333; text-align: center; }
        .box { background: #f5f5f5; border-radius: 8px; padding: 20px; margin: 15px 0; }
        .btn { background: #4CAF50; color: white; padding: 12px 30px; border: none; border-radius: 5px; 
               font-size: 16px; cursor: pointer; display: block; margin: 20px auto; }
        .btn:hover { background: #45a049; }
        .btn:disabled { background: #ccc; cursor: not-allowed; }
        input[type="file"] { font-size: 14px; }
        .format label { margin-right: 20px; font-size: 14px; }
        #status { text-align: center; margin: 15px 0; font-size: 14px; }
        .success { color: green; }
        .error { color: red; }
        .info { color: orange; }
    </style>
</head>
<body>
    <h1>接线总表生成器 V1.0</h1>
    
    <div class="box">
        <p><b>选择源文件：</b></p>
        <input type="file" id="fileInput" accept=".xlsm,.xlsx,.xls">
    </div>
    
    <div class="box">
        <p><b>输出格式：</b></p>
        <div class="format">
            <label><input type="radio" name="format" value="xlsx" checked> xlsx (Excel 2007+)</label>
            <label><input type="radio" name="format" value="xls"> xls (WPS兼容)</label>
        </div>
    </div>
    
    <button class="btn" id="generateBtn" onclick="generate()">生成接线总表</button>
    
    <div id="status"></div>

    <script>
        async function generate() {
            const fileInput = document.getElementById('fileInput');
            const format = document.querySelector('input[name="format"]:checked').value;
            const status = document.getElementById('status');
            const btn = document.getElementById('generateBtn');
            
            if (!fileInput.files.length) {
                status.innerHTML = '<span class="error">请先选择Excel文件！</span>';
                return;
            }
            
            const file = fileInput.files[0];
            const formData = new FormData();
            formData.append('file', file);
            formData.append('format', format);
            
            btn.disabled = true;
            status.innerHTML = '<span class="info">正在生成接线总表，请稍候...</span>';
            
            try {
                const response = await fetch('/generate', { method: 'POST', body: formData });
                const result = await response.json();
                
                if (result.success) {
                    status.innerHTML = '<span class="success">生成完成！共 ' + result.count + ' 行数据<br>文件: ' + result.path + '</span>';
                } else {
                    status.innerHTML = '<span class="error">生成失败: ' + result.error + '</span>';
                }
            } catch (e) {
                status.innerHTML = '<span class="error">请求失败: ' + e.message + '</span>';
            }
            
            btn.disabled = false;
        }
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/generate', methods=['POST'])
def generate():
    try:
        file = request.files['file']
        fmt = request.form.get('format', 'xlsx')
        
        # 保存上传的文件
        upload_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(upload_dir, file.filename)
        file.save(file_path)
        
        # 生成接线总表
        generator = WiringTableGenerator(file_path)
        output_path = os.path.join(upload_dir, f"接线总表.{fmt}")
        count = generator.generate(output_path)
        
        # 删除上传的临时文件
        try:
            os.remove(file_path)
        except:
            pass
        
        return jsonify({
            'success': True,
            'count': count,
            'path': output_path
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })

def open_browser():
    """延迟打开浏览器"""
    import time
    time.sleep(1)
    webbrowser.open('http://127.0.0.1:5000')

if __name__ == '__main__':
    print("=" * 50)
    print("  接线总表生成器 V1.0")
    print("  浏览器将自动打开，如未打开请访问:")
    print("  http://127.0.0.1:5000")
    print("  按 Ctrl+C 关闭程序")
    print("=" * 50)
    
    # 在新线程中打开浏览器
    threading.Thread(target=open_browser, daemon=True).start()
    
    app.run(host='127.0.0.1', port=5000, debug=False)
