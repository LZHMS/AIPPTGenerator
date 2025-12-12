/**
 * AI PPT Generator - 主要JavaScript文件
 */

// 全局变量
let currentSlide = 0;
let totalSlides = 0;
let selectedTheme = 'business';
let pptData = null;
let downloadUrl = null;

// 主题数据
const themes = {
    business: {
        name: "商务蓝",
        colors: {
            primary_color: "#2E86AB",
            secondary_color: "#1A365D",
            accent_color: "#F18F01",
            background_color: "#FFFFFF",
            text_color: "#333333",
            title_color: "#1A1A2E"
        }
    },
    tech: {
        name: "科技紫",
        colors: {
            primary_color: "#667eea",
            secondary_color: "#764ba2",
            accent_color: "#00D9FF",
            background_color: "#0F0F23",
            text_color: "#E0E0E0",
            title_color: "#FFFFFF"
        }
    },
    nature: {
        name: "自然绿",
        colors: {
            primary_color: "#2D6A4F",
            secondary_color: "#40916C",
            accent_color: "#95D5B2",
            background_color: "#FFFFFF",
            text_color: "#333333",
            title_color: "#1B4332"
        }
    },
    warm: {
        name: "温暖橙",
        colors: {
            primary_color: "#E85D04",
            secondary_color: "#DC2F02",
            accent_color: "#FFBA08",
            background_color: "#FFFBF5",
            text_color: "#333333",
            title_color: "#370617"
        }
    },
    minimal: {
        name: "极简黑白",
        colors: {
            primary_color: "#333333",
            secondary_color: "#666666",
            accent_color: "#E63946",
            background_color: "#FFFFFF",
            text_color: "#333333",
            title_color: "#000000"
        }
    },
    ocean: {
        name: "海洋蓝",
        colors: {
            primary_color: "#0077B6",
            secondary_color: "#023E8A",
            accent_color: "#48CAE4",
            background_color: "#CAF0F8",
            text_color: "#03045E",
            title_color: "#023E8A"
        }
    }
};

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', function() {
    initThemeSelector();
    initEventListeners();
    loadWorkflowDiagram();
    loadLLMInfo();
});

/**
 * 加载 LLM 信息
 */
async function loadLLMInfo() {
    try {
        const response = await fetch('/api/llm-info');
        const data = await response.json();
        const modelName = data.model_id || 'Unknown Model';
        document.getElementById('llmModelName').textContent = `Model: ${modelName}`;
    } catch (error) {
        console.error('Failed to load LLM info:', error);
        document.getElementById('llmModelName').textContent = 'Powered by LangGraph';
    }
}

/**
 * 初始化主题选择器
 */
function initThemeSelector() {
    const container = document.getElementById('themeSelector');
    container.innerHTML = '';
    
    for (const [key, theme] of Object.entries(themes)) {
        const option = document.createElement('div');
        option.className = 'theme-option' + (key === selectedTheme ? ' selected' : '');
        option.dataset.theme = key;
        
        option.innerHTML = `
            <div class="color-preview">
                <div class="color-dot" style="background-color: ${theme.colors.primary_color}"></div>
                <div class="color-dot" style="background-color: ${theme.colors.secondary_color}"></div>
                <div class="color-dot" style="background-color: ${theme.colors.accent_color}"></div>
            </div>
            <div class="theme-name">${theme.name}</div>
        `;
        
        option.addEventListener('click', () => selectTheme(key));
        container.appendChild(option);
    }
}

/**
 * 选择主题
 */
function selectTheme(themeKey) {
    selectedTheme = themeKey;
    
    // 更新选中状态
    document.querySelectorAll('.theme-option').forEach(opt => {
        opt.classList.toggle('selected', opt.dataset.theme === themeKey);
    });
}

/**
 * 初始化事件监听器
 */
function initEventListeners() {
    // 生成按钮
    document.getElementById('generateBtn').addEventListener('click', generatePPT);
    
    // 页数滑块
    const numSlidesInput = document.getElementById('numSlides');
    const numSlidesValue = document.getElementById('numSlidesValue');
    if (numSlidesInput && numSlidesValue) {
        numSlidesInput.addEventListener('input', function() {
            numSlidesValue.textContent = this.value + ' 页';
        });
    }
    
    // 幻灯片导航
    document.getElementById('prevSlide').addEventListener('click', () => navigateSlide(-1));
    document.getElementById('nextSlide').addEventListener('click', () => navigateSlide(1));
    
    // 下载按钮
    document.getElementById('downloadBtn').addEventListener('click', downloadPPT);
    
    // 键盘导航
    document.addEventListener('keydown', function(e) {
        if (totalSlides > 0) {
            if (e.key === 'ArrowLeft') navigateSlide(-1);
            if (e.key === 'ArrowRight') navigateSlide(1);
        }
    });
}

/**
 * 加载工作流图
 */
async function loadWorkflowDiagram() {
    try {
        const response = await fetch('/api/workflow/graph');
        const data = await response.json();
        
        if (data.mermaid) {
            const container = document.getElementById('workflowDiagram');
            container.innerHTML = `<div class="mermaid">${data.mermaid}</div>`;
            mermaid.init(undefined, '.mermaid');
        }
    } catch (error) {
        console.error('Failed to load workflow diagram:', error);
        document.getElementById('workflowDiagram').innerHTML = `
            <div class="text-center text-muted">
                <i class="bi bi-exclamation-triangle"></i>
                工作流图加载失败
            </div>
        `;
    }
}

/**
 * 生成PPT
 */
async function generatePPT() {
    const topic = document.getElementById('topicInput').value.trim();
    const numSlidesInput = document.getElementById('numSlides');
    const numSlides = numSlidesInput ? parseInt(numSlidesInput.value) : 6;
    
    if (!topic) {
        alert('请输入PPT主题！');
        return;
    }
    
    // 显示进度条，隐藏预览
    showProgress();
    hidePreview();
    
    // 禁用生成按钮
    const btn = document.getElementById('generateBtn');
    btn.disabled = true;
    btn.innerHTML = '<i class="bi bi-hourglass-split me-2 spin"></i>生成中...';
    
    // 设置超时控制器
    const controller = new AbortController();
    const timeoutId = setTimeout(() => {
        controller.abort();
        addLogItem('请求超时，请重试或减少页数', 'error');
    }, 300000); // 5分钟超时
    
    try {
        // 使用fetch + POST方式
        const response = await fetch('/api/generate/stream', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                topic: topic,
                theme: selectedTheme,
                num_slides: numSlides
            }),
            signal: controller.signal
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';
        let lastActivityTime = Date.now();
        
        // 心跳检测 - 如果60秒没有数据，显示警告
        const heartbeatCheck = setInterval(() => {
            if (Date.now() - lastActivityTime > 60000) {
                addLogItem('正在等待服务器响应...', 'warning');
            }
        }, 30000);
        
        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            
            lastActivityTime = Date.now();
            buffer += decoder.decode(value, { stream: true });
            
            // 处理SSE数据
            const lines = buffer.split('\n');
            buffer = lines.pop() || '';
            
            for (const line of lines) {
                if (line.startsWith('data: ')) {
                    try {
                        const data = JSON.parse(line.slice(6));
                        handleStreamData(data);
                    } catch (e) {
                        console.error('Parse error:', e);
                    }
                }
            }
        }
        
        clearInterval(heartbeatCheck);
        clearTimeout(timeoutId);
        
    } catch (error) {
        console.error('Generation error:', error);
        if (error.name === 'AbortError') {
            addLogItem('生成超时，请尝试减少页数后重试', 'error');
        } else {
            addLogItem(`生成失败: ${error.message}`, 'error');
        }
    } finally {
        clearTimeout(timeoutId);
        // 恢复按钮状态
        btn.disabled = false;
        btn.innerHTML = '<i class="bi bi-magic me-2"></i>开始生成';
    }
}

/**
 * 处理流式数据
 */
function handleStreamData(data) {
    switch (data.type) {
        case 'start':
            addLogItem(data.message);
            updateProgress(5);
            break;
            
        case 'progress':
            addLogItem(data.message, 'success');
            // 根据步骤更新进度
            const progressMap = {
                'search_resources': 20,
                'generate_theme_style': 30,
                'generate_color_scheme': 40,
                'generate_content_outline': 55,
                'design_slide_layouts': 70,
                'generate_detailed_content': 85,
                'assemble_ppt_data': 95
            };
            if (progressMap[data.step]) {
                updateProgress(progressMap[data.step]);
            }
            break;
            
        case 'complete':
            updateProgress(100);
            addLogItem('✓ ' + data.message, 'success');
            
            // 保存数据
            pptData = data.ppt_data;
            downloadUrl = data.download_url;
            
            // 显示预览
            setTimeout(() => {
                renderPreview(data.ppt_data);
                showDetails(data.ppt_data);
                document.getElementById('downloadBtn').style.display = 'inline-block';
            }, 500);
            break;
            
        case 'error':
            addLogItem(data.message, 'error');
            break;
    }
}

/**
 * 显示进度区域
 */
function showProgress() {
    const container = document.getElementById('progressContainer');
    container.style.display = 'block';
    document.getElementById('progressLog').innerHTML = '';
    updateProgress(0);
}

/**
 * 更新进度条
 */
function updateProgress(percent) {
    const bar = document.getElementById('progressBar');
    bar.style.width = percent + '%';
    bar.textContent = percent + '%';
}

/**
 * 添加日志项
 */
function addLogItem(message, type = '') {
    const log = document.getElementById('progressLog');
    const item = document.createElement('div');
    item.className = 'log-item' + (type ? ' ' + type : '');
    item.innerHTML = `<i class="bi bi-${type === 'error' ? 'x-circle' : type === 'success' ? 'check-circle' : 'arrow-right'}"></i> ${message}`;
    log.appendChild(item);
    log.scrollTop = log.scrollHeight;
}

/**
 * 隐藏预览
 */
function hidePreview() {
    document.getElementById('previewPlaceholder').style.display = 'flex';
    document.getElementById('pptPreview').style.display = 'none';
    document.getElementById('pptDetails').style.display = 'none';
    document.getElementById('downloadBtn').style.display = 'none';
}

/**
 * 渲染PPT预览
 */
function renderPreview(data) {
    document.getElementById('previewPlaceholder').style.display = 'none';
    document.getElementById('pptPreview').style.display = 'block';
    
    const container = document.getElementById('slidesContainer');
    container.innerHTML = '';
    
    const slides = data.slides || [];
    const colors = data.color_scheme || themes[selectedTheme].colors;
    
    totalSlides = slides.length;
    currentSlide = 0;
    
    slides.forEach((slide, index) => {
        const slideEl = document.createElement('div');
        slideEl.className = 'slide' + (index === 0 ? ' active' : '');
        slideEl.id = `slide-${index}`;
        
        // 根据类型设置样式
        switch (slide.slide_type) {
            case 'title':
                slideEl.classList.add('slide-title');
                slideEl.style.backgroundColor = colors.background_color;
                slideEl.innerHTML = `
                    <h1 style="color: ${colors.title_color}">${slide.title}</h1>
                    <div class="subtitle" style="color: ${colors.text_color}">${slide.subtitle || (slide.content && slide.content[0]) || ''}</div>
                    <div style="width: 100px; height: 4px; background: ${colors.primary_color}; margin-top: 20px;"></div>
                `;
                break;
                
            case 'summary':
                slideEl.classList.add('slide-summary');
                slideEl.style.background = `linear-gradient(135deg, ${colors.primary_color}, ${colors.secondary_color})`;
                slideEl.innerHTML = `
                    <h2>${slide.title}</h2>
                    <div class="summary-points">${(slide.content || []).join(' • ')}</div>
                `;
                break;
                
            default:
                slideEl.classList.add('slide-content');
                slideEl.style.backgroundColor = colors.background_color;
                
                let contentHTML = `
                    <h2 style="color: ${colors.primary_color}; border-color: ${colors.primary_color}">${slide.title}</h2>
                    <ul>
                `;
                
                (slide.content || []).forEach(item => {
                    contentHTML += `<li style="color: ${colors.text_color}">${item}</li>`;
                });
                
                contentHTML += '</ul>';
                slideEl.innerHTML = contentHTML;
        }
        
        container.appendChild(slideEl);
    });
    
    updateSlideIndicator();
}

/**
 * 导航幻灯片
 */
function navigateSlide(direction) {
    if (totalSlides === 0) return;
    
    // 隐藏当前幻灯片
    document.getElementById(`slide-${currentSlide}`).classList.remove('active');
    
    // 计算新索引
    currentSlide += direction;
    if (currentSlide < 0) currentSlide = totalSlides - 1;
    if (currentSlide >= totalSlides) currentSlide = 0;
    
    // 显示新幻灯片
    document.getElementById(`slide-${currentSlide}`).classList.add('active');
    
    updateSlideIndicator();
}

/**
 * 更新幻灯片指示器
 */
function updateSlideIndicator() {
    document.getElementById('slideIndicator').textContent = `${currentSlide + 1} / ${totalSlides}`;
}

/**
 * 显示详细信息
 */
function showDetails(data) {
    document.getElementById('pptDetails').style.display = 'block';
    
    // 主题风格
    const themeStyle = data.theme_style || {};
    document.getElementById('themeStyleInfo').textContent = 
        themeStyle.style_name || themes[selectedTheme].name;
    
    // 配色方案
    const colors = data.color_scheme || themes[selectedTheme].colors;
    const colorContainer = document.getElementById('colorSchemeInfo');
    colorContainer.innerHTML = '';
    
    for (const [key, value] of Object.entries(colors)) {
        if (value && value.startsWith('#')) {
            const block = document.createElement('div');
            block.className = 'color-block';
            block.style.backgroundColor = value;
            block.title = `${key}: ${value}`;
            colorContainer.appendChild(block);
        }
    }
    
    // 大纲
    const outlineList = document.getElementById('outlineInfo');
    outlineList.innerHTML = '';
    
    (data.slides || []).forEach((slide, index) => {
        const li = document.createElement('li');
        li.innerHTML = `<span class="slide-num">${index + 1}</span>${slide.title}`;
        outlineList.appendChild(li);
    });
}

/**
 * 下载PPT
 */
function downloadPPT() {
    if (downloadUrl) {
        window.location.href = downloadUrl;
    }
}
