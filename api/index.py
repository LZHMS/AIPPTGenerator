"""
Vercel Serverless Function 入口
"""
import os
import sys

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from web.app import app

# Vercel 需要的入口点
def handler(request, response):
    return app(request, response)

# 导出 Flask app 供 Vercel 使用
app = app
