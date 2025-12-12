"""
AI PPT Generator - 启动脚本
运行此脚本启动Web应用
"""

import os
import sys
import webbrowser
from threading import Timer

# 添加项目根目录到路径
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_root)

def open_browser():
    """在浏览器中打开应用"""
    webbrowser.open('http://127.0.0.1:5000')

if __name__ == '__main__':
    # 切换到web目录
    web_dir = os.path.join(project_root, 'web')
    os.chdir(web_dir)
    
    # 导入Flask应用
    from web.app import app
    
    print("=" * 60)
    print("  AI PPT Generator - 智能PPT生成器")
    print("=" * 60)
    print(f"  版本: 1.0.0")
    print(f"  访问地址: http://127.0.0.1:5000")
    print("=" * 60)
    print("\n提示：按 Ctrl+C 停止服务\n")
    
    # 3秒后自动打开浏览器
    Timer(3.0, open_browser).start()
    
    # 启动Flask应用
    app.run(debug=True, host='0.0.0.0', port=5000, use_reloader=False)
