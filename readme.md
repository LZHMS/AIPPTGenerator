---
title: AI PPT Generator
emoji: 📊
colorFrom: blue
colorTo: purple
sdk: docker
pinned: false
license: mit
app_port: 7860
---

# AI PPT Generator - 智能PPT生成器

一个基于 LangGraph 的 AI 驱动的 PPT 自动生成工具，支持多种主题风格，并提供网页端交互界面。

## 功能特点

- 🎯 **智能主题生成**: 根据输入主题自动生成合适的 PPT 风格
- 🎨 **多种配色方案**: 支持商务蓝、科技紫、自然绿等6种预设主题
- 🤖 **AI 内容生成**: 使用 LLM 直接生成高质量内容（无需额外搜索 API）
- 📊 **智能布局设计**: AI 自动设计幻灯片布局（15种布局类型）
- 🌐 **Web 界面**: 简洁易用的网页端操作界面
- 📥 **一键下载**: 生成的 PPT 可直接下载

## 工作流程

使用 LangGraph 构建的工作流包含以下节点：

```
┌──────────────────┐
│      START       │
└────────┬─────────┘
         │
    ┌────┴────┐
    ▼         ▼
┌─────────┐ ┌─────────────────┐
│ 搜索资源 │ │ 生成主题风格    │
└────┬────┘ └────────┬────────┘
     │               │
     │               ▼
     │        ┌─────────────────┐
     │        │ 生成配色方案    │
     │        └────────┬────────┘
     │               │
     ▼               │
┌─────────────────┐  │
│ 生成内容大纲    │  │
└────┬────────────┘  │
     │               │
     ├───────────────┤
     ▼               ▼
┌─────────────────┐ ┌─────────────────┐
│ 设计幻灯片布局  │ │ 生成详细内容    │
└────────┬────────┘ └────────┬────────┘
         │                   │
         └─────────┬─────────┘
                   ▼
          ┌───────────────┐
          │ 组装PPT数据   │
          └───────┬───────┘
                  ▼
          ┌───────────────┐
          │     END       │
          └───────────────┘
```

## 安装

1. 克隆项目后，进入目录：
```bash
cd ppt_generator
```

2. 安装依赖：
```bash
pip install -r requirements.txt
```

## 运行

### 方式一：直接运行启动脚本
```bash
python run.py
```

### 方式二：运行 Flask 应用
```bash
cd web
python app.py
```

启动后访问：http://127.0.0.1:5000

## 使用方法

1. 打开浏览器访问 http://127.0.0.1:5000
2. 在输入框中输入 PPT 主题（如："人工智能的发展与应用"）
3. 选择喜欢的主题风格（商务蓝、科技紫、自然绿等）
4. 点击"开始生成"按钮
5. 等待生成完成，预览 PPT
6. 点击"下载PPT"按钮获取文件

## 项目结构

```
ppt_generator/
├── __init__.py           # 包初始化
├── requirements.txt      # 依赖文件
├── run.py                # 启动脚本
├── README.md             # 说明文档
│
├── core/                 # 核心模块
│   ├── __init__.py
│   ├── ppt_agent.py      # LangGraph 工作流 Agent
│   └── ppt_builder.py    # PPT 构建器（python-pptx）
│
└── web/                  # Web 应用
    ├── app.py            # Flask 应用
    ├── output/           # 生成的 PPT 输出目录
    ├── templates/        # HTML 模板
    │   └── index.html
    └── static/           # 静态资源
        ├── css/
        │   └── style.css
        └── js/
            └── main.js
```

## 主题风格

| 主题名称 | 描述 | 适用场景 |
|---------|------|---------|
| 商务蓝 | 专业、稳重的蓝色系 | 商业报告、工作汇报 |
| 科技紫 | 现代、前卫的紫色渐变 | 科技产品、技术分享 |
| 自然绿 | 清新、环保的绿色系 | 环保主题、健康相关 |
| 温暖橙 | 活力、热情的橙色系 | 营销推广、活动策划 |
| 极简黑白 | 简约、大气的黑白配色 | 学术报告、正式场合 |
| 海洋蓝 | 清爽、宁静的蓝色调 | 旅游、文化主题 |

## API 配置

只需要一个 API Key！

### 环境变量配置

```bash
# Moonshot AI API（必需）
MOONSHOT_API_KEY=your-moonshot-api-key

# 可选：自定义模型配置
LLM_MODEL_ID=kimi-k2-thinking-turbo
LLM_BASE_URL=https://api.moonshot.cn/v1
```

获取 API Key：https://platform.moonshot.cn/

## 部署到 Hugging Face Spaces

### 步骤 1：创建 Space

1. 访问 https://huggingface.co/spaces
2. 点击 **Create new Space**
3. 填写：
   - **Space name**: `ai-ppt-generator`
   - **SDK**: 选择 **Docker**
4. 点击 **Create Space**

### 步骤 2：配置 Secrets

1. 进入 Space 的 **Settings** 页面
2. 找到 **Repository secrets**
3. 添加：`MOONSHOT_API_KEY` = 你的 API Key

### 步骤 3：上传代码

**方式一：Git 推送**
```bash
git remote add hf https://huggingface.co/spaces/你的用户名/ai-ppt-generator
git push hf main
```

**方式二：网页上传**
- 在 Space 的 **Files** 页面直接上传所有文件

### 步骤 4：等待构建

Hugging Face 会自动构建 Docker 镜像并部署，约 3-5 分钟完成。

## 本地运行

```bash
# 设置环境变量
export MOONSHOT_API_KEY=your-api-key

# 安装依赖
pip install -r requirements.txt

# 启动应用
python run.py
```

访问：http://127.0.0.1:5000

## 技术栈

- **LangGraph**: 工作流编排
- **LangChain**: LLM 集成
- **python-pptx**: PPT 文件生成
- **Flask**: Web 框架
- **Bootstrap 5**: 前端UI

## 注意事项

1. 首次运行需要联网下载依赖
2. 生成 PPT 需要调用 LLM API，请确保 API Key 配置正确
3. 生成的 PPT 文件保存在 `web/output/` 目录

## License

MIT License
