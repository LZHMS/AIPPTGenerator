"""
AI PPT Generator - Web Application
Flask Web 应用程序
"""

from flask import Flask, render_template, request, jsonify, send_file, Response
from flask_cors import CORS
import os
import sys
import json
import uuid
import time
from datetime import datetime

# 添加父目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.ppt_agent import generate_ppt_data, generate_ppt_data_stream, build_ppt_workflow, LLM_CONFIG
from core.ppt_builder import create_ppt_from_data, THEME_PRESETS, apply_theme_preset

app = Flask(__name__, 
            template_folder='templates',
            static_folder='static')
CORS(app)

# 配置
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.route('/')
def index():
    """主页"""
    return render_template('index.html', themes=THEME_PRESETS)


@app.route('/api/llm-info', methods=['GET'])
def get_llm_info():
    """获取当前 LLM 配置信息"""
    return jsonify({
        "model_id": LLM_CONFIG.get("model_id", "unknown"),
        "base_url": LLM_CONFIG.get("base_url", "unknown")
    })


@app.route('/api/themes', methods=['GET'])
def get_themes():
    """获取可用的主题列表"""
    themes = []
    for key, value in THEME_PRESETS.items():
        themes.append({
            "id": key,
            "name": value["name"],
            "colors": value["color_scheme"]
        })
    return jsonify({"themes": themes})


@app.route('/api/generate', methods=['POST'])
def generate_ppt():
    """生成PPT（非流式）"""
    try:
        data = request.get_json()
        topic = data.get('topic', '')
        theme = data.get('theme', 'business')
        num_slides = data.get('num_slides', 6)
        
        # 验证页数范围
        try:
            num_slides = int(num_slides)
            num_slides = max(4, min(20, num_slides))
        except (TypeError, ValueError):
            num_slides = 6
        
        if not topic:
            return jsonify({"error": "请输入PPT主题"}), 400
        
        # 生成PPT数据
        ppt_data = generate_ppt_data(topic, num_slides=num_slides)
        
        # 应用主题
        if theme in THEME_PRESETS:
            ppt_data = apply_theme_preset(ppt_data, theme)
        
        # 生成文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"ppt_{timestamp}_{uuid.uuid4().hex[:8]}.pptx"
        filepath = os.path.join(OUTPUT_DIR, filename)
        
        # 创建PPT文件
        create_ppt_from_data(ppt_data, filepath)
        
        return jsonify({
            "success": True,
            "message": "PPT生成成功",
            "filename": filename,
            "download_url": f"/api/download/{filename}",
            "ppt_data": ppt_data
        })
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/api/generate/stream', methods=['POST'])
def generate_ppt_stream():
    """流式生成PPT，返回Server-Sent Events"""
    data = request.get_json()
    topic = data.get('topic', '')
    theme = data.get('theme', 'business')
    num_slides = data.get('num_slides', 6)
    
    # 验证页数范围
    try:
        num_slides = int(num_slides)
        num_slides = max(4, min(20, num_slides))
    except (TypeError, ValueError):
        num_slides = 6
    
    if not topic:
        return jsonify({"error": "请输入PPT主题"}), 400
    
    def generate():
        try:
            # 发送开始事件
            yield f"data: {json.dumps({'type': 'start', 'message': '开始生成PPT...'})}\n\n"
            
            final_ppt_data = None
            last_heartbeat = time.time()
            
            # 流式生成
            for step, status, node_output in generate_ppt_data_stream(topic, num_slides=num_slides):
                progress_data = {
                    'type': 'progress',
                    'step': step,
                    'status': status,
                    'message': f"[{step}] {status}"
                }
                yield f"data: {json.dumps(progress_data, ensure_ascii=False)}\n\n"
                
                # 获取最终数据
                if 'ppt_data' in node_output and node_output['ppt_data']:
                    final_ppt_data = node_output['ppt_data']
                
                # 发送心跳保持连接
                current_time = time.time()
                if current_time - last_heartbeat > 15:
                    yield f"data: {json.dumps({'type': 'heartbeat', 'message': '处理中...'})}\n\n"
                    last_heartbeat = current_time
                
                time.sleep(0.1)  # 小延迟以便前端能看到进度
            
            if final_ppt_data:
                # 应用主题
                if theme in THEME_PRESETS:
                    final_ppt_data = apply_theme_preset(final_ppt_data, theme)
                
                # 生成文件
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"ppt_{timestamp}_{uuid.uuid4().hex[:8]}.pptx"
                filepath = os.path.join(OUTPUT_DIR, filename)
                
                create_ppt_from_data(final_ppt_data, filepath)
                
                # 发送完成事件
                complete_data = {
                    'type': 'complete',
                    'success': True,
                    'message': 'PPT生成完成！',
                    'filename': filename,
                    'download_url': f"/api/download/{filename}",
                    'ppt_data': final_ppt_data
                }
                yield f"data: {json.dumps(complete_data, ensure_ascii=False)}\n\n"
            else:
                yield f"data: {json.dumps({'type': 'error', 'message': '生成失败，未能获取PPT数据'})}\n\n"
                
        except Exception as e:
            error_data = {
                'type': 'error',
                'message': f"生成出错: {str(e)}"
            }
            yield f"data: {json.dumps(error_data, ensure_ascii=False)}\n\n"
    
    response = Response(generate(), mimetype='text/event-stream')
    response.headers['Cache-Control'] = 'no-cache'
    response.headers['X-Accel-Buffering'] = 'no'  # 禁用 Nginx 缓冲
    return response


@app.route('/api/download/<filename>')
def download_ppt(filename):
    """下载PPT文件"""
    filepath = os.path.join(OUTPUT_DIR, filename)
    
    if not os.path.exists(filepath):
        return jsonify({"error": "文件不存在"}), 404
    
    return send_file(
        filepath,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )


@app.route('/api/preview/<filename>')
def preview_ppt(filename):
    """获取PPT预览数据"""
    filepath = os.path.join(OUTPUT_DIR, filename.replace('.pptx', '.json'))
    
    # 尝试读取对应的JSON数据
    if os.path.exists(filepath):
        with open(filepath, 'r', encoding='utf-8') as f:
            return jsonify(json.load(f))
    
    return jsonify({"error": "预览数据不存在"}), 404


@app.route('/api/workflow/graph')
def get_workflow_graph():
    """获取工作流图表数据"""
    try:
        workflow = build_ppt_workflow()
        # 获取mermaid图表
        mermaid_code = workflow.get_graph().draw_mermaid()
        return jsonify({
            "mermaid": mermaid_code
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    print("=" * 50)
    print("AI PPT Generator Web Application")
    print("=" * 50)
    print(f"访问地址: http://127.0.0.1:5000")
    print(f"输出目录: {OUTPUT_DIR}")
    print("=" * 50)
    
    app.run(debug=True, host='0.0.0.0', port=5000)
