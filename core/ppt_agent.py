"""
AI PPT Generator - Core Agent Module
使用 LangGraph 构建 PPT 生成工作流
"""

from typing import TypedDict, List, Dict, Any, Optional, Annotated
from langgraph.graph import StateGraph, START, END
from langchain_openai import ChatOpenAI
import json
import os


# LLM 配置 - 优先从环境变量读取
LLM_CONFIG = {
    "api_key": os.environ.get("MOONSHOT_API_KEY", "your-api-key-here"),
    "model_id": os.environ.get("LLM_MODEL_ID", "kimi-k2-thinking-turbo"),
    "base_url": os.environ.get("LLM_BASE_URL", "https://api.moonshot.cn/v1")
}


def get_llm():
    """获取 LLM 实例"""
    return ChatOpenAI(
        model=LLM_CONFIG["model_id"],
        api_key=LLM_CONFIG["api_key"],
        base_url=LLM_CONFIG["base_url"],
        streaming=False,
    )


def merge_status(left: str, right: str) -> str:
    """合并状态信息"""
    if not left:
        return right
    if not right:
        return left
    return f"{left} | {right}"


# PPT 状态定义 - 使用 Annotated 处理并发更新
class PPTState(TypedDict):
    topic: str                                              # 用户输入的主题
    num_slides: int                                         # 指定页数
    search_results: List[Dict]                              # 搜索结果
    theme_style: Dict[str, Any]                             # 主题风格
    color_scheme: Dict[str, str]                            # 配色方案
    content_outline: List[Dict]                             # 内容大纲
    slide_layouts: List[Dict]                               # 幻灯片布局
    generated_content: List[Dict]                           # 生成的内容
    ppt_data: Dict[str, Any]                                # 最终 PPT 数据
    status: Annotated[str, merge_status]                    # 当前状态（支持并发更新）
    error: Optional[str]                                    # 错误信息


# 可用的布局类型
LAYOUT_TYPES = [
    "title",              # 标题页
    "content",            # 普通内容页
    "bullet_points",      # 要点列表
    "two_column",         # 两栏布局
    "three_column",       # 三栏布局
    "image_left",         # 左图右文
    "image_right",        # 右图左文
    "quote",              # 引用/名言页
    "statistics",         # 数据统计页
    "timeline",           # 时间线布局
    "comparison",         # 对比布局
    "icons_grid",         # 图标网格
    "big_number",         # 大数字展示
    "process_flow",       # 流程图
    "summary",            # 总结页
]


def get_layout_suggestions(num_slides: int) -> List[str]:
    """根据页数推荐布局组合"""
    if num_slides <= 4:
        return ["title", "content", "bullet_points", "summary"]
    elif num_slides <= 6:
        return ["title", "content", "bullet_points", "two_column", "statistics", "summary"]
    elif num_slides <= 8:
        return ["title", "content", "bullet_points", "two_column", "image_left", 
                "statistics", "timeline", "summary"]
    elif num_slides <= 10:
        return ["title", "content", "bullet_points", "two_column", "three_column",
                "image_left", "image_right", "statistics", "process_flow", "summary"]
    else:
        return ["title", "content", "bullet_points", "two_column", "three_column",
                "image_left", "image_right", "quote", "statistics", "timeline",
                "comparison", "icons_grid", "big_number", "process_flow", "summary"]


def generate_default_outline(topic: str, num_slides: int) -> List[Dict]:
    """生成默认大纲"""
    layouts = get_layout_suggestions(num_slides)
    outline = []
    
    # 第一页始终是标题页
    outline.append({
        "slide_number": 1, 
        "slide_type": "title", 
        "title": topic, 
        "key_points": ["探索与发现"], 
        "notes": "开场介绍"
    })
    
    # 中间页面
    middle_layouts = [l for l in layouts if l not in ["title", "summary"]]
    content_templates = [
        ("概述", ["核心概念", "主要特点", "发展历程"], "content"),
        ("核心要点", ["要点一：基础原理", "要点二：关键技术", "要点三：实践方法"], "bullet_points"),
        ("优势与特点", ["高效性", "可扩展性", "易用性"], "two_column"),
        ("应用场景", ["场景一", "场景二", "场景三"], "three_column"),
        ("数据分析", ["关键数据1", "关键数据2", "趋势分析"], "statistics"),
        ("发展历程", ["起步阶段", "发展阶段", "成熟阶段"], "timeline"),
        ("对比分析", ["传统方式", "创新方式"], "comparison"),
        ("核心功能", ["功能一", "功能二", "功能三", "功能四"], "icons_grid"),
        ("关键指标", ["核心数据展示"], "big_number"),
        ("实施流程", ["步骤1", "步骤2", "步骤3", "步骤4"], "process_flow"),
        ("案例展示", ["案例背景", "实施方案", "取得成效"], "image_left"),
        ("深度解析", ["详细分析", "专家观点"], "image_right"),
        ("名言启示", ["引用相关名言或观点"], "quote"),
    ]
    
    for i in range(1, num_slides - 1):
        if i - 1 < len(content_templates):
            title, points, layout = content_templates[i - 1]
        else:
            title = f"第{i}部分"
            points = ["内容要点"]
            layout = middle_layouts[(i - 1) % len(middle_layouts)] if middle_layouts else "content"
        
        outline.append({
            "slide_number": i + 1,
            "slide_type": layout,
            "title": title,
            "key_points": points,
            "notes": f"第{i + 1}页内容"
        })
    
    # 最后一页是总结页
    outline.append({
        "slide_number": num_slides,
        "slide_type": "summary",
        "title": "总结与展望",
        "key_points": ["回顾要点", "未来展望", "行动建议"],
        "notes": "结束语"
    })
    
    return outline[:num_slides]


# 节点函数
def search_resources(state: PPTState) -> Dict:
    """使用 LLM 生成相关资料（无需外部搜索 API）"""
    topic = state["topic"]
    num_slides = state.get("num_slides", 6)
    llm = get_llm()
    
    try:
        # 使用 LLM 生成关于主题的详细资料
        prompt = f"""作为一个知识专家，请为以下主题生成详细的背景资料和信息，用于制作一个 {num_slides} 页的 PPT。

主题：{topic}

请生成 {min(num_slides, 8)} 条相关资料，每条资料包含一个关键点。
以 JSON 数组格式返回，每个元素包含：
- "title": 资料标题/关键点名称
- "content": 详细内容描述（100-200字）
- "category": 类别（如：概述、特点、应用、趋势、案例等）

只返回 JSON 数组，不要其他内容。"""

        response = llm.invoke(prompt)
        content = response.content.strip()
        
        # 提取 JSON
        if "```json" in content:
            content = content.split("```json")[1].split("```")[0].strip()
        elif "```" in content:
            content = content.split("```")[1].split("```")[0].strip()
        
        results = json.loads(content)
        
        # 标准化格式
        parsed_results = []
        for r in results:
            parsed_results.append({
                "title": r.get("title", ""),
                "content": r.get("content", ""),
                "category": r.get("category", "概述")
            })
        
        return {
            "search_results": parsed_results,
            "status": f"资料生成完成（{len(parsed_results)}条）"
        }
    except Exception as e:
        # 返回基础资料
        return {
            "search_results": [
                {"title": f"{topic}概述", "content": f"关于{topic}的基础介绍", "category": "概述"},
                {"title": "核心特点", "content": f"{topic}的主要特点和优势", "category": "特点"},
                {"title": "应用场景", "content": f"{topic}的典型应用场景", "category": "应用"},
                {"title": "发展趋势", "content": f"{topic}的未来发展方向", "category": "趋势"},
            ],
            "status": "资料生成完成（基础模式）"
        }


def generate_theme_style(state: PPTState) -> Dict:
    """生成主题风格节点"""
    topic = state["topic"]
    llm = get_llm()
    
    prompt = f"""
    为主题为"{topic}"的PPT设计一个合适的主题风格。
    
    请以JSON格式返回，包含以下字段：
    {{
        "style_name": "风格名称（如：简约商务、科技未来、清新自然等）",
        "font_family": "推荐字体（如：微软雅黑、Arial等）",
        "title_font_size": "标题字号（数字）",
        "body_font_size": "正文字号（数字）",
        "design_elements": ["设计元素1", "设计元素2"],
        "mood": "整体氛围描述"
    }}
    
    只返回JSON，不要其他内容。
    """
    
    try:
        response = llm.invoke(prompt)
        content = response.content.strip()
        
        # 尝试提取JSON
        if "```json" in content:
            content = content.split("```json")[1].split("```")[0]
        elif "```" in content:
            content = content.split("```")[1].split("```")[0]
        
        theme_style = json.loads(content)
        
        return {
            "theme_style": theme_style,
            "status": "主题风格生成完成"
        }
    except Exception as e:
        # 默认主题风格
        return {
            "theme_style": {
                "style_name": "简约商务",
                "font_family": "微软雅黑",
                "title_font_size": 44,
                "body_font_size": 24,
                "design_elements": ["简洁线条", "留白设计"],
                "mood": "专业、清晰、易读"
            },
            "status": "主题风格生成完成（使用默认）"
        }


def generate_color_scheme(state: PPTState) -> Dict:
    """生成配色方案节点"""
    topic = state["topic"]
    theme_style = state.get("theme_style", {})
    llm = get_llm()
    
    style_name = theme_style.get("style_name", "简约商务")
    
    prompt = f"""
    为主题为"{topic}"、风格为"{style_name}"的PPT设计一套配色方案。
    
    请以JSON格式返回，包含以下字段（颜色使用十六进制格式）：
    {{
        "primary_color": "#主色调",
        "secondary_color": "#辅助色",
        "accent_color": "#强调色",
        "background_color": "#背景色",
        "text_color": "#正文文字颜色",
        "title_color": "#标题颜色",
        "gradient_start": "#渐变起始色",
        "gradient_end": "#渐变结束色"
    }}
    
    只返回JSON，不要其他内容。
    """
    
    try:
        response = llm.invoke(prompt)
        content = response.content.strip()
        
        if "```json" in content:
            content = content.split("```json")[1].split("```")[0]
        elif "```" in content:
            content = content.split("```")[1].split("```")[0]
        
        color_scheme = json.loads(content)
        
        return {
            "color_scheme": color_scheme,
            "status": "配色方案生成完成"
        }
    except Exception as e:
        # 默认配色方案
        return {
            "color_scheme": {
                "primary_color": "#2E86AB",
                "secondary_color": "#A23B72",
                "accent_color": "#F18F01",
                "background_color": "#FFFFFF",
                "text_color": "#333333",
                "title_color": "#1A1A2E",
                "gradient_start": "#667eea",
                "gradient_end": "#764ba2"
            },
            "status": "配色方案生成完成（使用默认）"
        }


def generate_content_outline(state: PPTState) -> Dict:
    """生成内容大纲节点"""
    topic = state["topic"]
    num_slides = state.get("num_slides", 6)
    search_results = state.get("search_results", [])
    llm = get_llm()
    
    # 整理搜索结果
    search_context = ""
    # 根据页数决定使用多少搜索结果
    results_to_use = min(len(search_results), max(3, num_slides // 2))
    for idx, result in enumerate(search_results[:results_to_use], 1):
        search_context += f"\n资料{idx}: {result.get('title', '')}\n{result.get('content', '')[:600]}\n"
    
    # 根据页数推荐布局组合
    layout_suggestions = get_layout_suggestions(num_slides)
    
    prompt = f"""
    为主题为"{topic}"的PPT生成一份内容大纲，要求生成正好 {num_slides} 页。
    
    参考资料：
    {search_context if search_context else "（无参考资料，请根据通用知识生成）"}
    
    可用的幻灯片布局类型（请尽量多样化使用）：
    - title: 标题页（第1页必须使用）
    - content: 普通内容页，带标题和段落文字
    - bullet_points: 要点列表，适合罗列多个要点
    - two_column: 两栏布局，适合对比或并列内容
    - three_column: 三栏布局，适合展示多个并列概念
    - image_left: 左图右文，适合图文结合
    - image_right: 右图左文，适合图文结合
    - quote: 引用页，适合展示名言或重要观点
    - statistics: 数据统计页，适合展示数字和统计数据
    - timeline: 时间线布局，适合展示发展历程
    - comparison: 对比布局，适合对比两种方案或观点
    - icons_grid: 图标网格，适合展示多个功能或特点
    - big_number: 大数字展示，适合突出关键数据
    - process_flow: 流程图，适合展示步骤或流程
    - summary: 总结页（最后一页必须使用）
    
    推荐的布局组合：{layout_suggestions}
    
    请以JSON数组格式返回，每个元素包含：
    {{
        "slide_number": 页码数字（1到{num_slides}）,
        "slide_type": "布局类型",
        "title": "幻灯片标题",
        "key_points": ["要点1", "要点2", "要点3"],
        "notes": "演讲备注"
    }}
    
    重要要求：
    1. 必须生成正好 {num_slides} 页
    2. 第1页必须是 title 类型
    3. 第{num_slides}页必须是 summary 类型
    4. 中间页面请多样化使用不同的布局类型，不要全部使用 bullet_points 或 content
    5. 内容要专业、有深度，充分利用参考资料
    
    只返回JSON数组，不要其他内容。
    """
    
    try:
        response = llm.invoke(prompt)
        content = response.content.strip()
        
        if "```json" in content:
            content = content.split("```json")[1].split("```")[0]
        elif "```" in content:
            content = content.split("```")[1].split("```")[0]
        
        content_outline = json.loads(content)
        
        # 确保页数正确
        if len(content_outline) != num_slides:
            # 如果页数不对，使用默认大纲
            content_outline = generate_default_outline(topic, num_slides)
        
        return {
            "content_outline": content_outline,
            "status": f"内容大纲生成完成（{len(content_outline)}页）"
        }
    except Exception as e:
        # 使用默认大纲
        return {
            "content_outline": generate_default_outline(topic, num_slides),
            "status": "内容大纲生成完成（使用默认）"
        }


def design_slide_layouts(state: PPTState) -> Dict:
    """设计幻灯片布局节点"""
    content_outline = state.get("content_outline", [])
    theme_style = state.get("theme_style", {})
    color_scheme = state.get("color_scheme", {})
    llm = get_llm()
    
    prompt = f"""
    根据以下内容大纲，为每页PPT设计具体的布局。
    
    大纲：{json.dumps(content_outline, ensure_ascii=False)}
    主题风格：{json.dumps(theme_style, ensure_ascii=False)}
    配色方案：{json.dumps(color_scheme, ensure_ascii=False)}
    
    请以JSON数组格式返回每页的布局设计，每个元素包含：
    {{
        "slide_number": 页码,
        "layout_type": "布局类型",
        "elements": [
            {{
                "type": "元素类型（title/subtitle/text/bullet_list/image_placeholder/shape）",
                "position": {{"x": 左边距百分比, "y": 上边距百分比, "width": 宽度百分比, "height": 高度百分比}},
                "style": {{"font_size": 字号, "bold": 是否粗体, "align": "对齐方式"}}
            }}
        ]
    }}
    
    只返回JSON数组，不要其他内容。
    """
    
    try:
        response = llm.invoke(prompt)
        content = response.content.strip()
        
        if "```json" in content:
            content = content.split("```json")[1].split("```")[0]
        elif "```" in content:
            content = content.split("```")[1].split("```")[0]
        
        slide_layouts = json.loads(content)
        
        return {
            "slide_layouts": slide_layouts,
            "status": "布局设计完成"
        }
    except Exception as e:
        # 为每页生成默认布局
        default_layouts = []
        for slide in content_outline:
            layout = {
                "slide_number": slide.get("slide_number", 1),
                "layout_type": slide.get("slide_type", "content"),
                "elements": [
                    {
                        "type": "title",
                        "position": {"x": 5, "y": 5, "width": 90, "height": 15},
                        "style": {"font_size": 44, "bold": True, "align": "center"}
                    },
                    {
                        "type": "text" if slide.get("slide_type") != "bullet_points" else "bullet_list",
                        "position": {"x": 5, "y": 25, "width": 90, "height": 70},
                        "style": {"font_size": 24, "bold": False, "align": "left"}
                    }
                ]
            }
            default_layouts.append(layout)
        
        return {
            "slide_layouts": default_layouts,
            "status": "布局设计完成（使用默认）"
        }


def generate_detailed_content(state: PPTState) -> Dict:
    """生成详细内容节点"""
    topic = state["topic"]
    content_outline = state.get("content_outline", [])
    search_results = state.get("search_results", [])
    llm = get_llm()
    
    # 整理搜索结果作为参考
    search_context = ""
    for idx, result in enumerate(search_results[:5], 1):
        search_context += f"\n资料{idx}: {result.get('content', '')[:300]}\n"
    
    prompt = f"""
    为主题"{topic}"的PPT生成详细的每页内容。
    
    大纲：{json.dumps(content_outline, ensure_ascii=False)}
    
    参考资料：{search_context if search_context else "（无）"}
    
    请为每页生成详细内容，以JSON数组格式返回，每个元素包含：
    {{
        "slide_number": 页码,
        "title": "标题",
        "subtitle": "副标题（可选）",
        "content": ["内容段落1或要点1", "内容段落2或要点2", ...],
        "footer": "页脚文字（可选）"
    }}
    
    内容要求：
    1. 标题简洁有力
    2. 每页3-5个要点
    3. 每个要点不超过20个字
    4. 内容专业、准确
    
    只返回JSON数组，不要其他内容。
    """
    
    try:
        response = llm.invoke(prompt)
        content = response.content.strip()
        
        if "```json" in content:
            content = content.split("```json")[1].split("```")[0]
        elif "```" in content:
            content = content.split("```")[1].split("```")[0]
        
        generated_content = json.loads(content)
        
        return {
            "generated_content": generated_content,
            "status": "详细内容生成完成"
        }
    except Exception as e:
        # 使用大纲作为默认内容
        default_content = []
        for slide in content_outline:
            default_content.append({
                "slide_number": slide.get("slide_number", 1),
                "title": slide.get("title", ""),
                "subtitle": "",
                "content": slide.get("key_points", []),
                "footer": ""
            })
        
        return {
            "generated_content": default_content,
            "status": "详细内容生成完成（使用大纲）"
        }


def assemble_ppt_data(state: PPTState) -> Dict:
    """组装最终PPT数据节点"""
    
    ppt_data = {
        "topic": state["topic"],
        "theme_style": state.get("theme_style", {}),
        "color_scheme": state.get("color_scheme", {}),
        "slides": []
    }
    
    generated_content = state.get("generated_content", [])
    slide_layouts = state.get("slide_layouts", [])
    content_outline = state.get("content_outline", [])
    
    # 组装每页数据
    for content in generated_content:
        slide_num = content.get("slide_number", 1)
        
        # 查找对应的布局
        layout = next(
            (l for l in slide_layouts if l.get("slide_number") == slide_num),
            {"layout_type": "content", "elements": []}
        )
        
        # 查找对应的大纲
        outline = next(
            (o for o in content_outline if o.get("slide_number") == slide_num),
            {"slide_type": "content"}
        )
        
        slide_data = {
            "slide_number": slide_num,
            "slide_type": outline.get("slide_type", "content"),
            "layout": layout,
            "title": content.get("title", ""),
            "subtitle": content.get("subtitle", ""),
            "content": content.get("content", []),
            "footer": content.get("footer", ""),
            "notes": outline.get("notes", "")
        }
        
        ppt_data["slides"].append(slide_data)
    
    # 按页码排序
    ppt_data["slides"].sort(key=lambda x: x["slide_number"])
    
    return {
        "ppt_data": ppt_data,
        "status": "PPT数据组装完成"
    }


def build_ppt_workflow():
    """构建 PPT 生成工作流"""
    
    # 创建状态图
    workflow = StateGraph(PPTState)
    
    # 添加节点
    workflow.add_node("search_resources", search_resources)
    workflow.add_node("generate_theme_style", generate_theme_style)
    workflow.add_node("generate_color_scheme", generate_color_scheme)
    workflow.add_node("generate_content_outline", generate_content_outline)
    workflow.add_node("design_slide_layouts", design_slide_layouts)
    workflow.add_node("generate_detailed_content", generate_detailed_content)
    workflow.add_node("assemble_ppt_data", assemble_ppt_data)
    
    # 定义边（工作流顺序）
    # 第一阶段：并行执行搜索和主题风格生成
    workflow.add_edge(START, "search_resources")
    workflow.add_edge(START, "generate_theme_style")
    
    # 第二阶段：配色方案依赖主题风格
    workflow.add_edge("generate_theme_style", "generate_color_scheme")
    
    # 第三阶段：内容大纲依赖搜索结果
    workflow.add_edge("search_resources", "generate_content_outline")
    
    # 第四阶段：布局设计依赖大纲和配色
    workflow.add_edge("generate_content_outline", "design_slide_layouts")
    workflow.add_edge("generate_color_scheme", "design_slide_layouts")
    
    # 第五阶段：详细内容依赖大纲和搜索结果
    workflow.add_edge("generate_content_outline", "generate_detailed_content")
    
    # 第六阶段：组装依赖所有前置步骤
    workflow.add_edge("design_slide_layouts", "assemble_ppt_data")
    workflow.add_edge("generate_detailed_content", "assemble_ppt_data")
    
    # 结束
    workflow.add_edge("assemble_ppt_data", END)
    
    return workflow.compile()


# 主执行函数
def generate_ppt_data(topic: str, num_slides: int = 6) -> Dict[str, Any]:
    """
    生成 PPT 数据
    
    Args:
        topic: PPT 主题
        num_slides: 页数（默认6页，范围4-20）
        
    Returns:
        PPT 数据字典
    """
    # 限制页数范围
    num_slides = max(4, min(20, num_slides))
    
    workflow = build_ppt_workflow()
    
    initial_state = {
        "topic": topic,
        "num_slides": num_slides,
        "search_results": [],
        "theme_style": {},
        "color_scheme": {},
        "content_outline": [],
        "slide_layouts": [],
        "generated_content": [],
        "ppt_data": {},
        "status": "",
        "error": None
    }
    
    # 执行工作流
    final_state = workflow.invoke(initial_state)
    
    return final_state["ppt_data"]


# 流式生成函数（用于进度显示）
def generate_ppt_data_stream(topic: str, num_slides: int = 6):
    """
    流式生成 PPT 数据，返回生成器
    
    Args:
        topic: PPT 主题
        num_slides: 页数（默认6页，范围4-20）
        
    Yields:
        (step_name, status, data) 元组
    """
    # 限制页数范围
    num_slides = max(4, min(20, num_slides))
    
    workflow = build_ppt_workflow()
    
    initial_state = {
        "topic": topic,
        "num_slides": num_slides,
        "search_results": [],
        "theme_style": {},
        "color_scheme": {},
        "content_outline": [],
        "slide_layouts": [],
        "generated_content": [],
        "ppt_data": {},
        "status": "",
        "error": None
    }
    
    # 使用 stream 方法获取中间状态
    for event in workflow.stream(initial_state):
        for node_name, node_output in event.items():
            status = node_output.get("status", "处理中...")
            yield (node_name, status, node_output)


if __name__ == "__main__":
    # 测试
    topic = "人工智能的未来发展"
    print(f"正在为主题 '{topic}' 生成PPT...")
    
    for step, status, data in generate_ppt_data_stream(topic):
        print(f"[{step}] {status}")
    
    print("\n生成完成！")
