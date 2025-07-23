import streamlit as st
import pandas as pd
import os
import glob
import math
import streamlit.components.v1 as components  # 添加components导入
import time
import re
import plotly.graph_objects as go  # 添加plotly导入，用于数据可视化
import openpyxl  # 明确导入openpyxl，用于Excel文件处理

# 设置页面配置
st.set_page_config(
    page_title="采购·拿货系统",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# 设置会话状态变量
if 'previous_page' not in st.session_state:
    st.session_state['previous_page'] = []
if 'current_page' not in st.session_state:
    st.session_state['current_page'] = 'home'
if 'previous_state' not in st.session_state:
    st.session_state['previous_state'] = []
if 'file_path' not in st.session_state:
    st.session_state['file_path'] = None
if 'data_loaded' not in st.session_state:
    st.session_state['data_loaded'] = False
if 'excel_data' not in st.session_state:
    st.session_state['excel_data'] = None
if 'tried_auto_load' not in st.session_state:
    st.session_state['tried_auto_load'] = False


# 格式化金额函数
def format_amount(value, unit="万元"):
    if pd.isna(value):
        return "0 " + unit
    if value >= 10000:
        return f"{value / 10000:.2f} 亿元"
    else:
        return f"{value:.2f} {unit}"


# 加载CSS样式
def load_css():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=SF+Pro+Display:wght@400;500;600;700&display=swap');
        @import url('https://fonts.googleapis.com/css2?family=SF+Pro+Text:wght@400;500&display=swap');
        @import url('https://fonts.googleapis.com/css2?family=SF+Mono&display=swap');

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
        }

        :root {
            --space-xxl: 4rem;
            --space-xl: 2.5rem;
            --space-lg: 1.8rem;
            --space-md: 1.2rem;
            --space-sm: 0.8rem;
            --space-xs: 0.4rem;

            --color-primary: #0A84FF;
            --color-secondary: #BF5AF2;
            --color-bg: #F5F5F7;  /* 米白色背景 */
            --color-surface: #FFFFFF;
            --color-card: rgba(255, 255, 255, 0.92);
            --color-text-primary: #1D1D1F;
            --color-text-secondary: #86868B;
            --color-text-tertiary: #8E8E93;
            --color-accent-red: #FF453A;
            --color-accent-blue: #0A84FF;
            --color-accent-green: #30D158;
            --color-accent-purple: #BF5AF2;
            --color-accent-yellow: #FFD60A;

            --glass-bg: rgba(255, 255, 255, 0.85);
            --glass-border: rgba(0, 0, 0, 0.08);
            --glass-highlight: rgba(0, 0, 0, 0.05);
        }

        /* 动画效果 */
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }

        @keyframes fadeInDown {
            from { opacity: 0; transform: translateY(-20px); }
            to { opacity: 1; transform: translateY(0); }
        }

        @keyframes pulse {
            0% { transform: scale(1); }
            50% { transform: scale(1.2); }
            100% { transform: scale(1); }
        }

        @keyframes shake {
            0% { transform: translateX(0); }
            10% { transform: translateX(-1px); }
            20% { transform: translateX(1px); }
            30% { transform: translateX(-1px); }
            40% { transform: translateX(1px); }
            50% { transform: translateX(0); }
            100% { transform: translateX(0); }
        }

        @keyframes float {
            0% { transform: translateY(0px); }
            50% { transform: translateY(-5px); }
            100% { transform: translateY(0px); }
        }

        .fade-in {
            animation: fadeIn 0.6s ease forwards;
        }

        body, .stApp {
            background: var(--color-bg);
            color: var(--color-text-primary);
            font-family: 'SF Pro Text', -apple-system, BlinkMacSystemFont, sans-serif;
            min-height: 100vh;
            background-attachment: fixed;
            background-image: radial-gradient(circle at 15% 50%, rgba(10, 132, 255, 0.05), transparent 40%),
                              radial-gradient(circle at 85% 30%, rgba(191, 90, 242, 0.05), transparent 40%);
        }

        .stDeployButton { display: none; }
        header[data-testid="stHeader"] { display: none; }
        .stMainBlockContainer { padding-top: 0; }

        /* 新导航栏样式 */
        .top-nav-container {
            background: var(--glass-bg);
            backdrop-filter: blur(20px) saturate(180%);
            -webkit-backdrop-filter: blur(20px) saturate(180%);
            border-radius: 50px;
            border: 0.5px solid var(--glass-border);
            box-shadow: 
                0 12px 30px rgba(0, 0, 0, 0.05),
                inset 0 0 0 1px rgba(255, 255, 255, 0.5);
            padding: 8px;
            margin: 20px auto 30px;
            max-width: 800px;
            display: flex;
            justify-content: center;
            align-items: center;
            transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.1);
            animation: fadeInDown 0.6s ease forwards;
            position: sticky;
            top: 10px;
            z-index: 100;
        }

        .top-nav-container:hover {
            transform: translateY(-2px);
            box-shadow: 
                0 15px 40px rgba(0, 0, 0, 0.08),
                inset 0 0 0 1px rgba(255, 255, 255, 0.7);
        }

        .nav-button {
            background: transparent;
            border: none;
            color: var(--color-text-primary);
            font-family: 'SF Pro Text', sans-serif;
            font-weight: 500;
            font-size: 1rem;
            padding: 12px 28px;
            margin: 0 5px;
            border-radius: 25px;
            transition: all 0.3s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
        }

        .nav-button:hover, .nav-button.active {
            background: rgba(255, 255, 255, 0.8);
            transform: translateY(-2px);
            box-shadow: 0 8px 20px rgba(0, 0, 0, 0.05);
        }

        .nav-button.active {
            background: linear-gradient(90deg, var(--color-primary), var(--color-secondary));
            color: white;
            box-shadow: 0 8px 20px rgba(10, 132, 255, 0.2);
        }

        .nav-button i {
            font-size: 1.2rem;
            margin-right: 8px;
            transition: transform 0.3s ease;
        }

        .nav-button:hover i {
            transform: scale(1.15);
        }

        /* 移除原来的分割线 */
        .stMainBlockContainer > hr {
            display: none !important;
        }

        /* 玻璃卡片 */
        .glass-card {
            background: var(--glass-bg);
            backdrop-filter: blur(20px) saturate(180%);
            -webkit-backdrop-filter: blur(20px) saturate(180%);
            border-radius: 18px;
            border: 0.5px solid var(--glass-border);
            box-shadow: 
                0 12px 30px rgba(0, 0, 0, 0.05),
                inset 0 0 0 1px rgba(255, 255, 255, 0.5);
            padding: var(--space-lg);
            transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.1);
            position: relative;
            overflow: hidden;
            margin-bottom: var(--space-md);
        }

        .glass-card:hover {
            transform: translateY(-6px);
            box-shadow: 
                0 20px 40px rgba(0, 0, 0, 0.08),
                inset 0 0 0 1px rgba(255, 255, 255, 0.7);
        }

        .glass-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 1px;
            background: linear-gradient(90deg, transparent, var(--glass-highlight), transparent);
        }

        /* 标题卡片 */
        .title-card {
            background: var(--glass-bg);
            backdrop-filter: blur(20px) saturate(180%);
            -webkit-backdrop-filter: blur(20px) saturate(180%);
            border-radius: 18px;
            border: 0.5px solid var(--glass-border);
            box-shadow: 
                0 12px 30px rgba(0, 0, 0, 0.05),
                inset 0 0 0 1px rgba(255, 255, 255, 0.5);
            padding: var(--space-sm) var(--space-lg);
            text-align: center;
            margin-bottom: var(--space-sm);
            margin-top: 0;
            animation: fadeIn 0.6s ease forwards;
        }

        /* 主标题 */
        .main-title {
            font-family: 'SF Pro Display', sans-serif;
            font-weight: 800;
            font-size: 3.6rem;
            background: linear-gradient(90deg, #64D2FF, #BF5AF2);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            letter-spacing: -0.5px;
            line-height: 1.1;
            margin-bottom: 0.8rem;
            text-shadow: 0px 0px 20px rgba(100, 210, 255, 0.3);
        }

        .main-subtitle {
            font-family: 'SF Pro Text', sans-serif;
            font-size: 1.7rem;
            font-weight: 500;
            color: var(--color-text-secondary);
            line-height: 1.6;
        }

        /* 上传区域 */
        .upload-area {
            background: rgba(255, 255, 255, 0.85);
            backdrop-filter: blur(10px);
            border: 1.5px dashed rgba(0, 0, 0, 0.1);
            border-radius: 18px;
            padding: var(--space-xl);
            text-align: center;
            transition: all 0.3s ease;
            color: var(--color-text-secondary);
        }

        .upload-area:hover {
            border-color: var(--color-primary);
            background: rgba(255, 255, 255, 0.92);
        }

        /* 标题样式 */
        .section-title {
            font-family: 'SF Pro Display', sans-serif;
            font-size: 2rem;
            font-weight: 700;
            margin: 0 0 var(--space-md);
            color: var(--color-text-primary);
            position: relative;
            padding-left: var(--space-md);
            letter-spacing: -0.5px;
        }

        .section-title:before {
            content: '';
            position: absolute;
            left: 0;
            top: 50%;
            transform: translateY(-50%);
            height: 70%;
            width: 4px;
            background: linear-gradient(to bottom, var(--color-primary), var(--color-secondary));
            border-radius: 4px;
        }

        .red-title:before {
            background: linear-gradient(to bottom, var(--color-accent-red), #FF375F);
        }

        .black-title:before {
            background: linear-gradient(to bottom, #8E8E93, #636366);
        }

        /* 按钮样式 */
        .stButton>button {
            background: linear-gradient(90deg, var(--color-primary), var(--color-secondary));
            color: white;
            border: none;
            border-radius: 12px;
            padding: 12px 25px;
            font-family: 'SF Pro Text', sans-serif;
            font-weight: 500;
            font-size: 1.1rem;
            transition: all 0.3s ease;
            box-shadow: 0 5px 15px rgba(10, 132, 255, 0.2);
        }

        .stButton>button:hover {
            transform: scale(1.03);
            box-shadow: 0 8px 20px rgba(191, 90, 242, 0.25);
        }

        /* 消息提示样式 */
        .stAlert {
            background: var(--glass-bg) !important;
            border-radius: 12px !important;
            backdrop-filter: blur(10px) !important;
            border: 0.5px solid rgba(0, 0, 0, 0.05) !important;
            box-shadow: 0 5px 15px rgba(0, 0, 0, 0.03) !important;
            margin-bottom: var(--space-md) !important;
        }

        /* 上传按钮 */
        .stFileUploader > div > div {
            background: var(--glass-bg) !important;
            border-radius: 18px !important;
            border: 0.5px solid rgba(0, 0, 0, 0.05) !important;
            padding: var(--space-md) !important;
            box-shadow: 0 5px 20px rgba(0, 0, 0, 0.03) !important;
            backdrop-filter: blur(10px) !important;
        }

        /* 进度条 */
        .stProgress > div > div > div {
            background: linear-gradient(90deg, var(--color-primary), var(--color-secondary)) !important;
        }

        /* 选择框 */
        .stSelectbox:not(div) {
            background: var(--glass-bg) !important;
            border-radius: 12px !important;
            border: 0.5px solid rgba(0, 0, 0, 0.05) !important;
        }

        /* 文本输入 */
        .stTextInput input {
            background: var(--glass-bg) !important;
            border-radius: 12px !important;
            border: 0.5px solid rgba(0, 0, 0, 0.05) !important;
            color: var(--color-text-primary) !important;
        }

        /* 重新设计的文件上传卡片 */
        .upload-card {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            padding: var(--space-xl) var(--space-lg);
            text-align: center;
            transition: all 0.4s ease;
            border: 2px dashed rgba(0, 0, 0, 0.08);
            background: linear-gradient(135deg, rgba(255,255,255,0.9), rgba(255,255,255,0.7));
        }

        .upload-card:hover {
            border-color: var(--color-primary);
            transform: translateY(-5px);
        }

        .upload-icon {
            width: 80px;
            height: 80px;
            display: flex;
            align-items: center;
            justify-content: center;
            background: linear-gradient(135deg, var(--color-primary), var(--color-secondary));
            border-radius: 50%;
            margin-bottom: var(--space-md);
            color: white;
            transition: all 0.3s ease;
        }

        .upload-card:hover .upload-icon {
            transform: scale(1.1);
        }

        .upload-title {
            font-family: 'SF Pro Display', sans-serif;
            font-weight: 700;
            font-size: 1.8rem;
            margin-bottom: var(--space-xs);
            color: var(--color-text-primary);
        }

        .upload-description {
            font-family: 'SF Pro Text', sans-serif;
            font-size: 1.1rem;
            color: var(--color-text-secondary);
            margin-bottom: var(--space-sm);
        }

        /* 文件状态卡片 */
        .file-status-card {
            display: flex;
            align-items: center;
            padding: var(--space-sm) var(--space-md);
        }

        .file-status-icon {
            width: 48px;
            height: 48px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-right: var(--space-md);
            flex-shrink: 0;
        }

        .status-success {
            color: var(--color-accent-green);
        }

        .status-info {
            color: var(--color-accent-blue);
        }

        .status-warning {
            color: var(--color-accent-yellow);
        }

        .file-status-content {
            flex-grow: 1;
        }

        .file-status-title {
            font-family: 'SF Pro Display', sans-serif;
            font-weight: 600;
            font-size: 1.2rem;
            margin-bottom: 2px;
        }

        .file-status-description {
            font-family: 'SF Pro Text', sans-serif;
            font-size: 0.95rem;
            color: var(--color-text-secondary);
        }

        /* 功能导航卡片样式 */
        .menu-card {
            padding: var(--space-lg);
        }

        .menu-header {
            display: flex;
            align-items: center;
            margin-bottom: var(--space-sm);
        }

        .menu-title-icon {
            display: flex;
            align-items: center;
            justify-content: center;
            width: 48px;
            height: 48px;
            background: linear-gradient(135deg, var(--color-primary), var(--color-secondary));
            border-radius: 12px;
            margin-right: var(--space-sm);
            box-shadow: 0 5px 15px rgba(10, 132, 255, 0.2);
        }

        .menu-title-icon svg {
            color: white;
        }

        .menu-title-text {
            font-family: 'SF Pro Display', sans-serif;
            font-weight: 700;
            font-size: 1.8rem;
            color: var(--color-text-primary);
            letter-spacing: -0.5px;
        }

        .menu-subtitle {
            font-family: 'SF Pro Text', sans-serif;
            font-size: 1rem;
            color: var(--color-text-secondary);
            margin-bottom: var(--space-md);
            margin-left: 60px;
        }

        .menu-divider {
            height: 1px;
            background: linear-gradient(to right, transparent, var(--glass-highlight), transparent);
            margin-bottom: var(--space-md);
        }

        .menu-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: var(--space-md);
            margin-top: var(--space-md);
        }

        .menu-item {
            position: relative;
            background: var(--glass-bg);
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
            border-radius: 16px;
            padding: var(--space-md);
            transition: all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.1);
            border: 1px solid rgba(255, 255, 255, 0.5);
            box-shadow: 0 5px 15px rgba(0, 0, 0, 0.05);
            overflow: hidden;
            cursor: pointer;
        }

        .menu-item:hover {
            transform: translateY(-5px) scale(1.02);
            box-shadow: 0 15px 30px rgba(0, 0, 0, 0.08);
        }

        .menu-item.primary .menu-icon {
            background: linear-gradient(135deg, #0A84FF, #0055c1);
        }

        .menu-item.success .menu-icon {
            background: linear-gradient(135deg, #30D158, #00923f);
        }

        .menu-item.info .menu-icon {
            background: linear-gradient(135deg, #64D2FF, #0090d0);
        }

        .menu-item.warning .menu-icon {
            background: linear-gradient(135deg, #FFD60A, #e0a800);
        }

        .menu-item.purple .menu-icon {
            background: linear-gradient(135deg, #BF5AF2, #9529ce);
        }

        .menu-icon {
            width: 56px;
            height: 56px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 12px;
            margin-bottom: var(--space-sm);
            color: white;
            transition: all 0.3s;
        }

        .menu-item:hover .menu-icon {
            transform: scale(1.1);
        }

        .menu-content {
            margin-left: var(--space-xs);
        }

        .menu-title {
            font-family: 'SF Pro Display', sans-serif;
            font-weight: 600;
            font-size: 1.3rem;
            margin-bottom: 4px;
            color: var(--color-text-primary);
        }

        .menu-description {
            font-family: 'SF Pro Text', sans-serif;
            font-size: 0.9rem;
            color: var(--color-text-secondary);
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }

        /* 卡片内容样式 */
        .card-content {
            padding: 0;
        }

        /* 调整标题样式 */
        .glass-card h3.section-title {
            margin: 0 0 var(--space-xs) 0;
            padding-top: 0;
            text-align: center;
            width: 100%;
        }

        /* 紧凑型卡片 */
        .compact-card {
            padding: var(--space-sm);
        }

        /* 新的卡片标题样式 */
        .card-title {
            text-align: center;
            margin: 8px 0 16px 0;
            padding: 8px 0;
            color: var(--color-text-primary);
            font-size: 1.2rem;
            font-weight: 600;
            width: 100%;
            border-bottom: 1px solid rgba(0,0,0,0.05);
        }

        /* 导航栏样式 */
        .nav-container {
            margin-bottom: 10px;
            padding: 0;
        }

        /* 导航按钮样式 */
        .stButton > button {
            background: var(--glass-bg);
            backdrop-filter: blur(20px) saturate(180%);
            -webkit-backdrop-filter: blur(20px) saturate(180%);
            border: 0.5px solid var(--glass-border);
            box-shadow: 0 5px 15px rgba(0, 0, 0, 0.05);
            color: var(--color-text-primary);
            font-family: 'SF Pro Text', sans-serif;
            font-weight: 500;
            font-size: 1rem;
            height: auto;
            padding: 8px 15px;
            border-radius: 50px;
            transition: all 0.3s ease;
            width: 100%;
            margin-right: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        /* 选项卡卡片样式 */
        .stTabs [data-baseweb="tab-list"] {
            gap: 16px;
            margin-bottom: 16px;
        }

        .stTabs [data-baseweb="tab"] {
            background-color: var(--glass-bg);
            border-radius: 16px;
            padding: 10px 20px;
            border: 0.5px solid var(--glass-border);
            box-shadow: 0 5px 15px rgba(0, 0, 0, 0.05);
            transition: all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.1);
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
            font-family: 'SF Pro Text', sans-serif;
            font-weight: 500;
            font-size: 1rem;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .stTabs [data-baseweb="tab"]:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 20px rgba(0, 0, 0, 0.1);
        }

        .stTabs [data-baseweb="tab"]:active {
            transform: scale(0.95);
        }

        .stTabs [data-baseweb="tab"][aria-selected="true"] {
            background: linear-gradient(135deg, var(--color-primary), var(--color-secondary));
            color: white;
            border: none;
            box-shadow: 0 8px 20px rgba(10, 132, 255, 0.2);
        }

        /* 隐藏选项卡下方的线 */
        .stTabs [data-baseweb="tab-highlight"] {
            display: none;
        }

        /* 选项卡内容区域样式 */
        .stTabs [data-baseweb="tab-panel"] {
            animation: fadeIn 0.5s ease forwards;
        }

        /* 为选项卡内容添加淡入动画 */
        @keyframes fadeIn {
            from {
                opacity: 0;
                transform: translateY(10px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
    </style>

    <script>
        // 动态光照效果
        document.addEventListener('DOMContentLoaded', function() {
            const cards = document.querySelectorAll('.glass-card');

            cards.forEach(card => {
                card.addEventListener('mousemove', function(e) {
                    const rect = this.getBoundingClientRect();
                    const x = e.clientX - rect.left;
                    const y = e.clientY - rect.top;

                    this.style.setProperty('--x', x + 'px');
                    this.style.setProperty('--y', y + 'px');
                });
            });
        });
    </script>
    """, unsafe_allow_html=True)


# 添加卡片组件函数
def create_glass_card(content, height=None, padding="1.5rem", margin="0 0 1.5rem 0", text_align="left"):
    """
    创建一个带有统一样式和悬停动画效果的玻璃卡片

    参数:
    content (str): 卡片内的HTML内容
    height (str, optional): 卡片高度，如"300px"
    padding (str, optional): 卡片内边距，默认"1.5rem"
    margin (str, optional): 卡片外边距，默认"0 0 1.5rem 0"
    text_align (str, optional): 文本对齐方式，默认"left"

    返回:
    str: 完整的卡片HTML
    """
    height_style = f"height: {height};" if height else ""

    card_html = f"""
    <div class="glass-card" style="padding: {padding}; margin: {margin}; border-radius: 18px; 
         background: var(--glass-bg); backdrop-filter: blur(20px) saturate(180%); 
         -webkit-backdrop-filter: blur(20px) saturate(180%); border: 0.5px solid var(--glass-border); 
         box-shadow: 0 12px 30px rgba(0, 0, 0, 0.05), inset 0 0 0 1px rgba(255, 255, 255, 0.5);
         text-align: {text_align}; transition: transform 0.3s ease, box-shadow 0.3s ease; 
         {height_style}">
        {content}
    </div>

    <style>
    .glass-card:hover {{
        transform: translateY(-5px);
        box-shadow: 0 15px 35px rgba(0, 0, 0, 0.1), inset 0 0 0 1px rgba(255, 255, 255, 0.6);
    }}
    </style>
    """
    return card_html


# 标题卡片组件
def title_card(title, subtitle=None):
    # 创建带有内联样式的卡片内容，确保样式直接应用
    content = f'''
    <style>
        .title-gradient {{
            font-family: 'SF Pro Display', sans-serif;
            font-weight: 800;
            font-size: 3.6rem;
            background: linear-gradient(90deg, #64D2FF, #BF5AF2);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            letter-spacing: -0.5px;
            line-height: 1.1;
            margin-bottom: 0.8rem;
            text-shadow: 0px 0px 20px rgba(100, 210, 255, 0.3);
        }}
        .subtitle-style {{
            font-family: 'SF Pro Text', sans-serif;
            font-size: 1.7rem;
            font-weight: 500;
            color: #86868B;
            line-height: 1.6;
            padding-bottom: 10px; /* 增加底部内边距，确保内容不被裁切 */
        }}
    </style>
    <div class="title-gradient">{title}</div>
    {f'<div class="subtitle-style">{subtitle}</div>' if subtitle else ''}
    '''

    # 使用create_glass_card创建玻璃效果卡片，增加内边距和卡片高度
    # 增加底部外边距，使卡片与下方内容保持更大距离
    card_html = create_glass_card(content, padding="2rem", margin="0 0 2.5rem 0", text_align="center")

    # 渲染卡片，增加高度以确保完整显示，包括底部圆角
    components.html(card_html, height=200)


# 自动检测Excel文件
def is_running_locally():
    """
    判断是否在本地环境运行
    在Streamlit Cloud上环境变量通常有STREAMLIT_SHARING或STREAMLIT_SERVER_URL
    """
    # 检测是否在Streamlit Cloud环境运行
    for env_var in ['STREAMLIT_SHARING', 'STREAMLIT_SERVER_URL', 'STREAMLIT_CLOUD']:
        if os.environ.get(env_var):
            return False

    # 如果没有检测到Cloud环境变量，就认为是在本地运行
    return True


def auto_detect_excel_file():
    """
    自动检测当前目录下的Excel文件
    只在本地环境中执行
    """
    # 在云端环境中不执行自动检测
    if not is_running_locally():
        return None

    try:
        # 支持不同的年月格式匹配：采购2023年_1月统计表.xlsx, 采购2023年_01月统计表.xlsx 等
        patterns = ["采购????年_?月统计表*.xlsx", "采购????年_??月统计表*.xlsx"]
        all_files = []

        # 遍历所有可能的模式
        for pattern in patterns:
            matched_files = glob.glob(pattern)
            if matched_files:  # 确保匹配到文件才添加
                all_files.extend(matched_files)

        if all_files:
            # 按文件修改时间排序，获取最新的文件
            latest_file = max(all_files, key=os.path.getctime)
            return latest_file

        return None
    except FileNotFoundError as e:
        st.error(f"文件未找到: {e}")
        return None
    except PermissionError as e:
        st.error(f"文件访问权限错误: {e}")
        return None
    except OSError as e:
        st.error(f"操作系统错误: {e}")
        return None
    except Exception as e:
        st.error(f"文件检测出错: {e}")
        return None


# 加载Excel数据
def load_excel_data(file_obj, target="main"):
    try:
        # 读取采购详情工作表（必需）
        details_df = None
        try:
            details_df = pd.read_excel(file_obj, sheet_name='采购详情', engine='openpyxl')
        except Exception as e:
            return None, "加载文件失败", f"无法读取'采购详情'工作表: {str(e)}"

        # 读取采购统计工作表（必需）
        stats_df = None
        try:
            stats_df = pd.read_excel(file_obj, sheet_name='采购统计', engine='openpyxl')
        except Exception as e:
            return None, "加载文件失败", f"无法读取'采购统计'工作表: {str(e)}"

        # 读取本月拿货统计工作表（必需）
        delivery_df = None
        try:
            delivery_df = pd.read_excel(file_obj, sheet_name='本月拿货统计', engine='openpyxl')
        except Exception as e:
            return None, "加载文件失败", f"无法读取'本月拿货统计'工作表: {str(e)}"

        # 保存数据
        excel_data = {
            '采购详情': details_df,
            '采购统计': stats_df,
            '本月拿货统计': delivery_df
        }

        # 仅当target为"main"时才保存到主会话状态
        if target == "main":
            st.session_state['excel_data'] = excel_data
            st.session_state['data_loaded'] = True

            # 如果是UploadedFile对象，保存文件名
            if hasattr(file_obj, 'name'):
                file_name = file_obj.name
                st.session_state['file_path'] = file_name
            else:
                # 如果是路径，保存路径
                st.session_state['file_path'] = file_obj
                file_name = os.path.basename(file_obj)
        else:
            # 仅获取文件名，不修改主会话状态
            if hasattr(file_obj, 'name'):
                file_name = file_obj.name
            else:
                file_name = os.path.basename(file_obj)

        # 返回数据和成功信息
        return excel_data, "加载文件成功", file_name
    except Exception as e:
        return None, "加载文件失败", f"读取文件时出错: {str(e)}"


# 显示导航栏
def show_navigation():
    # 添加导航栏样式
    st.markdown('''
    <style>
        /* 导航栏样式 */
        .nav-container {
            margin-bottom: 10px;
            padding: 0;
        }

        /* 导航按钮样式 */
        .stButton > button {
            background: var(--glass-bg);
            backdrop-filter: blur(20px) saturate(180%);
            -webkit-backdrop-filter: blur(20px) saturate(180%);
            border: 0.5px solid var(--glass-border);
            box-shadow: 0 5px 15px rgba(0, 0, 0, 0.05);
            color: var(--color-text-primary);
            font-family: 'SF Pro Text', sans-serif;
            font-weight: 500;
            font-size: 1rem;
            height: auto;
            padding: 8px 15px;
            border-radius: 50px;
            transition: all 0.3s ease;
            width: 100%;
            margin-right: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .stButton > button:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(0, 0, 0, 0.1);
            background: rgba(255, 255, 255, 0.9);
        }

        /* 图标 */
        .nav-icon {
            font-size: 1.3em;
            margin-right: 8px;
        }
    </style>
    ''', unsafe_allow_html=True)

    # 创建导航栏容器
    with st.container():
        st.markdown('<div class="nav-container">', unsafe_allow_html=True)

        # 创建三列布局
        col1, col2, col3 = st.columns(3)

        # 主页按钮
        with col1:
            if st.button("🏠 主页"):
                # 保存当前状态到撤销历史
                save_current_state()

                # 更新导航历史
                st.session_state['previous_page'].append(st.session_state['current_page'])
                st.session_state['current_page'] = 'home'
                st.rerun()

        # 返回按钮
        with col2:
            back_disabled = len(st.session_state['previous_page']) == 0
            if st.button("⬅️ 返回", disabled=back_disabled):
                # 保存当前状态到撤销历史
                save_current_state()

                # 检查是否在历月对比的子页面
                if st.session_state['current_page'] == 'history_compare' and st.session_state.get(
                        'history_compare_subpage') is not None:
                    # 从子页面返回到历月对比主页
                    st.session_state['history_compare_subpage'] = None
                else:
                    # 普通页面返回到上一级
                    st.session_state['current_page'] = st.session_state['previous_page'].pop()
                st.rerun()

        # 撤销按钮
        with col3:
            undo_disabled = len(st.session_state['previous_state']) == 0
            if st.button("↩️ 撤销", disabled=undo_disabled):
                prev_state = st.session_state['previous_state'].pop()
                for key, value in prev_state.items():
                    if key in st.session_state:
                        st.session_state[key] = value
                st.rerun()

        st.markdown('</div>', unsafe_allow_html=True)


# 显示主页
def show_home_page():
    # 主标题和副标题放在标题卡片中
    title_card("采购·拿货系统", "专业采购数据分析工具")

    # 添加一些额外的空间
    st.markdown('<div style="height:15px"></div>', unsafe_allow_html=True)

    # 创建左右分栏，左侧占1/3，右侧占2/3
    col1, col2 = st.columns([1, 2], gap="large")

    # 左侧文件上传区 - 使用新的卡片样式
    with col1:
        # 添加动态效果的上传卡片内容
        upload_content = '''
        <style>
            @keyframes pulseGlow {
                0% { box-shadow: 0 0 0 0 rgba(10, 132, 255, 0.2); }
                70% { box-shadow: 0 0 0 15px rgba(10, 132, 255, 0); }
                100% { box-shadow: 0 0 0 0 rgba(10, 132, 255, 0); }
            }

            @keyframes floatIcon {
                0% { transform: translateY(0px); }
                50% { transform: translateY(-8px); }
                100% { transform: translateY(0px); }
            }

            @keyframes gradientMove {
                0% { background-position: 0% 50%; }
                50% { background-position: 100% 50%; }
                100% { background-position: 0% 50%; }
            }

            .upload-icon-wrapper {
                position: relative;
                width: 90px;
                height: 90px;
                margin: 0 auto 15px auto;
                border-radius: 50%;
                background: linear-gradient(135deg, rgba(10, 132, 255, 0.1), rgba(191, 90, 242, 0.1));
                display: flex;
                align-items: center;
                justify-content: center;
                animation: pulseGlow 2s infinite cubic-bezier(0.66, 0, 0, 1), floatIcon 5s ease-in-out infinite;
            }

            .upload-icon {
                width: 70px;
                height: 70px;
                background: linear-gradient(135deg, #64D2FF, #BF5AF2);
                background-size: 200% 200%;
                animation: gradientMove 5s ease infinite;
                border-radius: 50%;
                display: flex;
                align-items: center;
                justify-content: center;
                box-shadow: 0 10px 20px rgba(10, 132, 255, 0.2);
                transition: transform 0.3s ease;
            }

            .upload-icon svg {
                filter: drop-shadow(0px 2px 3px rgba(0, 0, 0, 0.1));
            }

            .upload-title {
                font-family: 'SF Pro Display', sans-serif;
                font-weight: 600;
                font-size: 1.8rem;
                color: var(--color-text-primary);
                margin: 10px 0;
            }

            .upload-description {
                font-family: 'SF Pro Text', sans-serif;
                font-size: 1.1rem;
                color: var(--color-text-secondary);
                margin-bottom: 0.5rem;
                position: relative;
                display: inline-block;
            }

            .upload-description:after {
                content: '';
                position: absolute;
                width: 100%;
                height: 2px;
                bottom: -4px;
                left: 0;
                background: linear-gradient(90deg, #64D2FF, #BF5AF2);
                transform: scaleX(0);
                transform-origin: bottom right;
                transition: transform 0.5s cubic-bezier(0.86, 0, 0.07, 1);
            }

            .glass-card:hover .upload-description:after {
                transform: scaleX(1);
                transform-origin: bottom left;
            }

            .glass-card:hover .upload-icon {
                transform: scale(1.1);
            }
        </style>
        <div class="upload-icon-wrapper">
            <div class="upload-icon">
                <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                    <polyline points="17 8 12 3 7 8"></polyline>
                    <line x1="12" y1="3" x2="12" y2="15"></line>
                </svg>
            </div>
        </div>
        <h3 class="upload-title">数据上传</h3>
        <div class="upload-description">上传采购月度统计表</div>
        '''

        # 使用create_glass_card函数创建上传卡片
        upload_card_html = create_glass_card(upload_content, text_align="center", padding="2rem")
        components.html(upload_card_html, height=260)

        # 检查是否已经有数据加载
        already_loaded = st.session_state.get('data_loaded', False)

        # 本地环境自动检测Excel文件
        if not already_loaded and is_running_locally():
            # 检查会话状态中是否已经尝试过自动加载
            if not st.session_state.get('tried_auto_load', False):
                st.session_state['tried_auto_load'] = True

                # 尝试自动检测文件
                local_file = auto_detect_excel_file()
                if local_file:
                    with st.spinner(f"自动检测到文件 {os.path.basename(local_file)}，正在加载..."):
                        try:
                            # 加载检测到的文件
                            excel_data, success_msg, file_name = load_excel_data(local_file)
                            if excel_data is not None:  # 确保excel_data非空
                                # 使用新的卡片样式显示成功消息
                                auto_detect_content = f"""
                                <div style="display: flex; align-items: center; justify-content: center; padding: 0.5rem;">
                                    <span style="font-size: 1.5rem; margin-right: 0.5rem;">✅</span>
                                    <span style="font-size: 1rem; color: var(--color-text-success);">自动加载成功</span>
                                </div>
                                <div style="text-align: center; color: var(--color-text-secondary); margin-top: 0.5rem;">
                                    已加载: {file_name}
                                </div>
                                """
                                auto_detect_html = create_glass_card(auto_detect_content, text_align="center")
                                components.html(auto_detect_html, height=120)
                        except FileNotFoundError as e:
                            show_error_card(f"文件未找到: {str(e)}")
                        except PermissionError as e:
                            show_error_card(f"文件访问权限错误: {str(e)}")
                        except OSError as e:
                            show_error_card(f"操作系统错误: {str(e)}")
                        except Exception as e:
                            show_error_card(f"自动加载文件出错: {str(e)}")

        # 文件上传控件
        uploaded_file = st.file_uploader(
            "选择Excel文件",
            type=["xlsx"],
            help="请上传包含'采购详情'、'采购统计'和'本月拿货统计'工作表的Excel文件",
            label_visibility="collapsed"
        )

        # 处理上传的文件
        if uploaded_file is not None:
            with st.spinner("正在处理文件..."):
                try:
                    # 直接加载上传的文件数据
                    excel_data, success_msg, file_name = load_excel_data(uploaded_file)
                    if excel_data is None:
                        show_error_card(f"文件加载失败: {file_name}")
                    else:
                        # 使用新的卡片样式显示成功消息
                        success_content = f"""
                        <div style="display: flex; align-items: center; justify-content: center; padding: 0.5rem;">
                            <span style="font-size: 1.5rem; margin-right: 0.5rem;">✅</span>
                            <span style="font-size: 1rem; color: var(--color-text-success);">{success_msg}</span>
                        </div>
                        <div style="text-align: center; color: var(--color-text-secondary); margin-top: 0.5rem;">
                            已加载: {file_name}
                        </div>
                        """
                        success_html = create_glass_card(success_content, text_align="center")
                        components.html(success_html, height=120)
                except Exception as e:
                    show_error_card(f"处理文件时出错: {str(e)}")

    # 右侧功能菜单 - 使用新的卡片样式
    with col2:
        # 功能导航卡片的标题部分，添加动态效果
        menu_header = '''
        <style>
            @keyframes gradientText {
                0% { background-position: 0% 50%; }
                50% { background-position: 100% 50%; }
                100% { background-position: 0% 50%; }
            }

            @keyframes iconRotate {
                0% { transform: rotate(0deg); }
                25% { transform: rotate(90deg); }
                50% { transform: rotate(180deg); }
                75% { transform: rotate(270deg); }
                100% { transform: rotate(360deg); }
            }

            @keyframes typing {
                from { width: 0; }
                to { width: 100%; }
            }

            @keyframes pulseGlow {
                0% { box-shadow: 0 0 0 0 rgba(100, 210, 255, 0.4); }
                70% { box-shadow: 0 0 0 10px rgba(100, 210, 255, 0); }
                100% { box-shadow: 0 0 0 0 rgba(100, 210, 255, 0); }
            }

            @keyframes menuReveal {
                0% { transform: translateY(-10px); opacity: 0; }
                100% { transform: translateY(0); opacity: 1; }
            }

            .animated-menu-title {
                display: flex;
                align-items: center;
                animation: menuReveal 0.6s ease forwards;
            }

            .menu-title-icon {
                position: relative;
                width: 48px;
                height: 48px;
                background: linear-gradient(135deg, #64D2FF, #BF5AF2);
                background-size: 200% 200%;
                animation: gradientText 5s ease infinite;
                border-radius: 12px;
                display: flex;
                align-items: center;
                justify-content: center;
                margin-right: 15px;
                box-shadow: 0 8px 16px rgba(100, 210, 255, 0.2);
                animation: pulseGlow 2s infinite;
                transition: transform 0.3s ease;
            }

            .menu-title-icon svg {
                fill: white;
                animation: iconRotate 20s linear infinite;
            }

            .menu-title-text {
                font-family: 'SF Pro Display', sans-serif;
                font-size: 1.8rem;
                font-weight: 700;
                background: linear-gradient(90deg, #64D2FF, #BF5AF2);
                background-size: 200% 200%;
                animation: gradientText 5s ease infinite;
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                letter-spacing: -0.5px;
            }

            .menu-subtitle {
                font-family: 'SF Pro Text', sans-serif;
                font-size: 1.1rem;
                color: var(--color-text-secondary);
                margin: 8px 0 12px 63px;
                position: relative;
                display: inline-block;
                overflow: hidden;
                white-space: nowrap;
                border-right: 3px solid #64D2FF;
                animation: typing 3.5s steps(30, end) 0.5s 1 normal both, 
                           blink-caret 0.75s step-end infinite;
            }

            @keyframes blink-caret {
                from, to { border-color: transparent; }
                50% { border-color: #64D2FF; }
            }

            .menu-divider {
                height: 2px;
                background: linear-gradient(90deg, #64D2FF33, #BF5AF233, #64D2FF33);
                margin: 15px 0;
                border-radius: 2px;
                animation: gradientText 5s ease infinite;
                background-size: 200% 200%;
            }

            .menu-grid {
                animation: menuReveal 0.8s ease forwards;
            }
        </style>
        <div class="animated-menu-title">
            <div class="menu-title-icon">
                <svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M3 3h7v7H3z"></path>
                    <path d="M14 3h7v7h-7z"></path>
                    <path d="M14 14h7v7h-7z"></path>
                    <path d="M3 14h7v7H3z"></path>
                </svg>
            </div>
            <div class="menu-title-text">功能导航</div>
        </div>
        <div class="menu-subtitle">选择您需要的功能模块</div>
        <div class="menu-divider"></div>
        <div class="menu-grid">
        '''

        # 使用create_glass_card函数创建导航卡片头部
        menu_card_html = create_glass_card(menu_header, padding="1.5rem 1.5rem 0 1.5rem")
        components.html(menu_card_html, height=180)

        # 第一行按钮 - 3个
        first_row_items = [
            {
                "icon": '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"></path><path d="M13.73 21a2 2 0 0 1-3.46 0"></path></svg>',
                "title": "荣誉榜",
                "description": "查看采购团队荣誉榜",
                "key": "btn_honor",
                "page": "leaderboard",
                "color": "warning"
            },
            {
                "icon": '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"></path></svg>',
                "title": "采购详情",
                "description": "查看采购明细数据",
                "key": "btn_purchase_detail",
                "page": "purchase_detail",
                "color": "info"
            },
            {
                "icon": '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 12h-4l-3 9L9 3l-3 9H2"></path></svg>',
                "title": "采购统计",
                "description": "查看采购数据统计分析",
                "key": "btn_purchase_stats",
                "page": "purchase_stats",
                "color": "primary"
            }
        ]

        # 显示第一行按钮
        cols1 = st.columns(3)
        for i, item in enumerate(first_row_items):
            with cols1[i]:
                # 创建一个简单的按钮，但应用自定义CSS使其看起来像卡片
                if st.button(
                        f"**{item['title']}**\n\n{item['description']}",
                        key=item["key"],
                        use_container_width=True
                ):
                    st.session_state['previous_page'].append(st.session_state['current_page'])
                    st.session_state['current_page'] = item["page"]
                    st.rerun()

                # 添加自定义CSS，为按钮添加图标和样式
                st.markdown(f"""
                <style>
                    /* 为按钮 {item['key']} 添加自定义样式 */
                    [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"] {{
                        background: var(--glass-bg);
                        border-radius: 16px;
                        padding: 1.5rem 1rem;
                        height: auto;
                        min-height: 150px;
                        display: flex;
                        flex-direction: column;
                        align-items: flex-start;
                        text-align: left;
                        font-family: 'SF Pro Text', sans-serif;
                        transition: all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.1);
                        border: 1px solid rgba(255, 255, 255, 0.5);
                        position: relative;
                        overflow: hidden;
                    }}

                    [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"]:hover {{
                        transform: translateY(-5px) scale(1.02);
                        box-shadow: 0 15px 30px rgba(0, 0, 0, 0.08);
                    }}

                    [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"]:before {{
                        content: '';
                        position: absolute;
                        top: 1rem;
                        left: 1rem;
                        width: 48px;
                        height: 48px;
                        border-radius: 12px;
                        background: linear-gradient(135deg, var(--color-primary), var(--color-secondary));
                        background-image: url("data:image/svg+xml;charset=utf8,{item['icon'].replace('"', "'")}");
                        background-position: center;
                        background-repeat: no-repeat;
                        background-size: 24px;
                        margin-bottom: 1rem;
                    }}

                    /* 调整按钮的文字部分 */
                    [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"] p {{
                        padding-left: 70px;
                        padding-top: 5px;
                    }}

                    /* 使按钮的第一行文字(标题)粗体且大一点 */
                    [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"] p strong {{
                        font-size: 1.3rem;
                        font-weight: 600;
                        display: block;
                        margin-bottom: 0.5rem;
                    }}

                    /* 设置按钮背景颜色 */
                    [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"]:before {{
                        background: linear-gradient(135deg, 
                            {get_color_for_item(item['color'])});
                    }}
                </style>
                """, unsafe_allow_html=True)

        # 第二行按钮 - 2个居中
        second_row_items = [
            {
                "icon": '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2l9 4.9V17L12 22l-9-4.9V7z"/><circle cx="12" cy="12" r="4"/></svg>',
                "title": "拿货统计",
                "description": "查看拿货数据统计分析",
                "key": "btn_delivery_stats",
                "page": "delivery_stats",
                "color": "success"
            },
            {
                "icon": '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="20" x2="12" y2="10"></line><line x1="18" y1="20" x2="18" y2="4"></line><line x1="6" y1="20" x2="6" y2="16"></line></svg>',
                "title": "历月对比",
                "description": "历月采购拿货数据对比",
                "key": "btn_history_compare",
                "page": "history_compare",
                "color": "purple"
            }
        ]

        # 创建一个1:2:2:1比例的列布局，中间两列放按钮
        col_left, col_mid1, col_mid2, col_right = st.columns([1, 2, 2, 1])

        with col_mid1:
            item = second_row_items[0]
            if st.button(
                    f"**{item['title']}**\n\n{item['description']}",
                    key=item["key"],
                    use_container_width=True
            ):
                st.session_state['previous_page'].append(st.session_state['current_page'])
                st.session_state['current_page'] = item["page"]
                st.rerun()

            # 添加自定义CSS
            st.markdown(f"""
            <style>
                /* 为按钮 {item['key']} 添加自定义样式 */
                [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"] {{
                    background: var(--glass-bg);
                    border-radius: 16px;
                    padding: 1.5rem 1rem;
                    height: auto;
                    min-height: 150px;
                    display: flex;
                    flex-direction: column;
                    align-items: flex-start;
                    text-align: left;
                    font-family: 'SF Pro Text', sans-serif;
                    transition: all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.1);
                    border: 1px solid rgba(255, 255, 255, 0.5);
                    position: relative;
                    overflow: hidden;
                }}

                [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"]:hover {{
                    transform: translateY(-5px) scale(1.02);
                    box-shadow: 0 15px 30px rgba(0, 0, 0, 0.08);
                }}

                [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"]:before {{
                    content: '';
                    position: absolute;
                    top: 1rem;
                    left: 1rem;
                    width: 48px;
                    height: 48px;
                    border-radius: 12px;
                    background: linear-gradient(135deg, var(--color-primary), var(--color-secondary));
                    background-image: url("data:image/svg+xml;charset=utf8,{item['icon'].replace('"', "'")}");
                    background-position: center;
                    background-repeat: no-repeat;
                    background-size: 24px;
                    margin-bottom: 1rem;
                }}

                /* 调整按钮的文字部分 */
                [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"] p {{
                    padding-left: 70px;
                    padding-top: 5px;
                }}

                /* 使按钮的第一行文字(标题)粗体且大一点 */
                [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"] p strong {{
                    font-size: 1.3rem;
                    font-weight: 600;
                    display: block;
                    margin-bottom: 0.5rem;
                }}

                /* 设置按钮背景颜色 */
                [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"]:before {{
                    background: linear-gradient(135deg, 
                        {get_color_for_item(item['color'])});
                }}
            </style>
            """, unsafe_allow_html=True)

        with col_mid2:
            item = second_row_items[1]
            if st.button(
                    f"**{item['title']}**\n\n{item['description']}",
                    key=item["key"],
                    use_container_width=True
            ):
                st.session_state['previous_page'].append(st.session_state['current_page'])
                st.session_state['current_page'] = item["page"]
                st.rerun()

            # 添加自定义CSS
            st.markdown(f"""
            <style>
                /* 为按钮 {item['key']} 添加自定义样式 */
                [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"] {{
                    background: var(--glass-bg);
                    border-radius: 16px;
                    padding: 1.5rem 1rem;
                    height: auto;
                    min-height: 150px;
                    display: flex;
                    flex-direction: column;
                    align-items: flex-start;
                    text-align: left;
                    font-family: 'SF Pro Text', sans-serif;
                    transition: all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.1);
                    border: 1px solid rgba(255, 255, 255, 0.5);
                    position: relative;
                    overflow: hidden;
                }}

                [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"]:hover {{
                    transform: translateY(-5px) scale(1.02);
                    box-shadow: 0 15px 30px rgba(0, 0, 0, 0.08);
                }}

                [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"]:before {{
                    content: '';
                    position: absolute;
                    top: 1rem;
                    left: 1rem;
                    width: 48px;
                    height: 48px;
                    border-radius: 12px;
                    background: linear-gradient(135deg, var(--color-primary), var(--color-secondary));
                    background-image: url("data:image/svg+xml;charset=utf8,{item['icon'].replace('"', "'")}");
                    background-position: center;
                    background-repeat: no-repeat;
                    background-size: 24px;
                    margin-bottom: 1rem;
                }}

                /* 调整按钮的文字部分 */
                [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"] p {{
                    padding-left: 70px;
                    padding-top: 5px;
                }}

                /* 使按钮的第一行文字(标题)粗体且大一点 */
                [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"] p strong {{
                    font-size: 1.3rem;
                    font-weight: 600;
                    display: block;
                    margin-bottom: 0.5rem;
                }}

                /* 设置按钮背景颜色 */
                [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"]:before {{
                    background: linear-gradient(135deg, 
                        {get_color_for_item(item['color'])});
                }}
            </style>
            """, unsafe_allow_html=True)

        st.markdown('</div></div>', unsafe_allow_html=True)


# 颜色辅助函数
def get_color_for_item(color_name):
    color_map = {
        "primary": "#0A84FF, #0055c1",
        "success": "#30D158, #00923f",
        "info": "#64D2FF, #0090d0",
        "warning": "#FFD60A, #e0a800",
        "purple": "#BF5AF2, #9529ce"
    }
    return color_map.get(color_name, "#0A84FF, #0055c1")  # 默认为蓝色


# 模糊处理金额函数（向上取整，两位有效数字）
def fuzzy_amount(amount_wan):
    if amount_wan < 10:  # 小于10万，向上取整到整数
        return math.ceil(amount_wan)
    elif amount_wan < 100:  # 10-100万，向上取整到10
        return math.ceil(amount_wan / 10) * 10
    elif amount_wan < 1000:  # 100-1000万，向上取整到10
        return math.ceil(amount_wan / 10) * 10
    else:  # 1000万以上，向上取整到100
        return math.ceil(amount_wan / 100) * 100


def show_warning_card(message, icon="⚠️"):
    """显示带有警告样式的卡片"""
    warning_content = f"""
    <div style="display: flex; align-items: center; justify-content: center; padding: 1rem;">
        <span style="font-size: 1.5rem; margin-right: 0.5rem;">{icon}</span>
        <span style="font-size: 1rem; color: var(--color-text-danger);">{message}</span>
    </div>
    """
    warning_html = create_glass_card(warning_content, text_align="center")
    components.html(warning_html, height=100)


def show_info_card(message, icon="ℹ️"):
    """显示带有信息样式的卡片"""
    info_content = f"""
    <div style="display: flex; align-items: center; justify-content: center; padding: 1rem;">
        <span style="font-size: 1.5rem; margin-right: 0.5rem;">{icon}</span>
        <span style="font-size: 1rem; color: var(--color-text-info);">{message}</span>
    </div>
    """
    info_html = create_glass_card(info_content, text_align="center")
    components.html(info_html, height=100)


def show_error_card(message, icon="❌"):
    """显示带有错误样式的卡片"""
    error_content = f"""
    <div style="display: flex; align-items: center; justify-content: center; padding: 1rem;">
        <span style="font-size: 1.5rem; margin-right: 0.5rem;">{icon}</span>
        <span style="font-size: 1rem; color: var(--color-text-danger);">{message}</span>
    </div>
    """
    error_html = create_glass_card(error_content, text_align="center")
    components.html(error_html, height=100)


# 显示荣誉榜页面
def show_leaderboard_page():
    title_card("采购荣誉榜")

    # 检查是否已加载数据
    if not st.session_state.get('data_loaded', False) or st.session_state['excel_data'] is None:
        show_warning_card("请先上传采购统计表数据")
        return

    # 获取采购统计数据和拿货统计数据
    excel_data = st.session_state['excel_data']
    stats_df = excel_data.get('采购统计')
    details_df = excel_data.get('采购详情')  # 添加采购详情数据
    delivery_df = excel_data.get('本月拿货统计')

    if stats_df is None:
        show_error_card("未能找到采购统计数据")
        return

    # 外部容器
    st.markdown('<div class="glass-card">', unsafe_allow_html=True)

    # 创建两行各两个卡片的布局
    row1_col1, row1_col2 = st.columns(2)
    row2_col1, row2_col2 = st.columns(2)

    # 1. 大单之王 - 大单采购金额最高者(从采购详情中获取)
    with row1_col1:
        # 检查是否有采购详情数据和"大单采购金额"列
        if details_df is not None and '大单采购金额' in details_df.columns:
            # 确定姓名列名
            detail_name_col = '姓名' if '姓名' in details_df.columns else '报价员'

            if detail_name_col in details_df.columns:
                # 筛选有大单采购金额的行
                large_orders = details_df[details_df['大单采购金额'].notna() & (details_df['大单采购金额'] != 0)]

                if not large_orders.empty:
                    # 按姓名分组，找出每个人的最大大单采购金额
                    grouped = large_orders.groupby(detail_name_col)['大单采购金额'].max().reset_index()
                    # 找出大单采购金额最高的人
                    max_big_order_person = grouped.loc[grouped['大单采购金额'].idxmax()]
                    max_big_order = max_big_order_person['大单采购金额']
                    top_person_name = max_big_order_person[detail_name_col]

                    # 显示大单之王卡片
                    st.markdown(f'''
                    <div class="glass-card" style="padding:1.2rem; text-align:center; border-radius:16px; box-shadow:0 10px 25px rgba(10, 132, 255, 0.18); border:1px solid rgba(10, 132, 255, 0.2); position:relative; overflow:hidden;">
                        <div style="position:absolute; top:-15px; right:-15px; width:80px; height:80px; background:linear-gradient(135deg, rgba(10, 132, 255, 0.2), transparent); border-radius:50%; z-index:1;"></div>
                        <div style="position:absolute; bottom:-20px; left:-20px; width:100px; height:100px; background:linear-gradient(135deg, transparent, rgba(10, 132, 255, 0.15)); border-radius:50%; z-index:1;"></div>
                        <h4 style="margin-bottom:1rem; color:var(--color-accent-blue); font-size:1.4rem; font-weight:700;">🏆 大单之王</h4>
                        <div style="display:flex; flex-direction:column; align-items:center; position:relative; z-index:2;">
                            <div style="width:90px; height:90px; border-radius:50%; background:linear-gradient(135deg, #0A84FF, #0055c1); display:flex; align-items:center; justify-content:center; margin-bottom:1rem; box-shadow:0 8px 20px rgba(10, 132, 255, 0.3); border:3px solid #ffffff;">
                                <span style="font-size:2.2rem; color:white;">👑</span>
                            </div>
                            <h3 style="margin:0.8rem 0; font-size:1.6rem; background:linear-gradient(90deg, #0055c1, #0A84FF); -webkit-background-clip:text; -webkit-text-fill-color:transparent;">{top_person_name}</h3>
                            <div style="background:rgba(10, 132, 255, 0.12); padding:0.5rem 1rem; border-radius:30px; margin-top:0.5rem;">
                                <p style="font-size:1.3rem; font-weight:700; color:#0A84FF; margin:0;">{fuzzy_amount(max_big_order / 10000)} 万元</p>
                            </div>
                        </div>
                    </div>
                    ''', unsafe_allow_html=True)
                else:
                    st.markdown('''
                    <div class="glass-card" style="padding:1rem; text-align:center;">
                        <h4 style="margin-bottom:0.8rem; color:var(--color-accent-blue);">🏆 大单之王</h4>
                        <div style="padding: 1rem;">
                            <p>没有找到大单采购数据</p>
                        </div>
                    </div>
                    ''', unsafe_allow_html=True)
            else:
                st.markdown('''
                <div class="glass-card" style="padding:1rem; text-align:center;">
                    <h4 style="margin-bottom:0.8rem; color:var(--color-accent-blue);">🏆 大单之王</h4>
                    <div style="padding: 1rem;">
                        <p>采购详情数据中缺少姓名列</p>
                    </div>
                </div>
                ''', unsafe_allow_html=True)
        else:
            st.markdown('''
            <div class="glass-card" style="padding:1rem; text-align:center;">
                <h4 style="margin-bottom:0.8rem; color:var(--color-accent-blue);">🏆 大单之王</h4>
                <div style="padding: 1rem;">
                    <p>未找到采购详情数据或数据中缺少大单采购金额列</p>
                </div>
            </div>
            ''', unsafe_allow_html=True)

    # 2. 量采之星 - 小单数最多者
    with row1_col2:
        # 确定姓名列名
        name_col = '姓名' if '姓名' in stats_df.columns else '报价员'

        # 检查是否有"小单数"列
        if '小单数' in stats_df.columns and name_col in stats_df.columns:
            # 排除"合计"行
            stats_no_total = stats_df[stats_df[name_col] != '合计'].copy()

            if not stats_no_total.empty:
                # 找出小单数最多的人
                max_small_orders = stats_no_total['小单数'].max()
                top_small_person = stats_no_total[stats_no_total['小单数'] == max_small_orders].iloc[0]

                # 显示量采之星卡片
                st.markdown(f'''
                <div class="glass-card" style="padding:1.2rem; text-align:center; border-radius:20px 8px 20px 8px; box-shadow:0 10px 25px rgba(48, 209, 88, 0.18); border:1px solid rgba(48, 209, 88, 0.2); position:relative; overflow:hidden;">
                    <div style="position:absolute; top:10px; left:10px; width:20px; height:20px; background:#30D158; opacity:0.2; border-radius:50%;"></div>
                    <div style="position:absolute; bottom:40px; right:20px; width:15px; height:15px; background:#30D158; opacity:0.15; border-radius:50%;"></div>
                    <div style="position:absolute; top:60px; right:40px; width:10px; height:10px; background:#30D158; opacity:0.1; border-radius:50%;"></div>
                    <h4 style="margin-bottom:1rem; color:var(--color-accent-green); font-size:1.4rem; font-weight:700;">
                        <span style="display:inline-block; animation:pulse 2s infinite;">🌟</span> 量采之星
                    </h4>
                    <div style="display:flex; flex-direction:column; align-items:center; position:relative; z-index:2;">
                        <div style="width:90px; height:90px; border-radius:20px; background:linear-gradient(135deg, #30D158, #00923f); display:flex; align-items:center; justify-content:center; margin-bottom:1rem; box-shadow:0 8px 20px rgba(48, 209, 88, 0.3); transform:rotate(5deg);">
                            <span style="font-size:2.2rem; color:white;">📊</span>
                        </div>
                        <h3 style="margin:0.8rem 0; font-size:1.6rem; color:#00923f;">{top_small_person[name_col]}</h3>
                        <div style="background:rgba(48, 209, 88, 0.12); padding:0.5rem 1rem; border-radius:30px; margin-top:0.5rem; position:relative; overflow:hidden;">
                            <div style="position:absolute; top:0; left:0; height:100%; width:{int(max_small_orders) / 50 * 100}%; max-width:100%; background:rgba(48, 209, 88, 0.1); z-index:1;"></div>
                            <p style="font-size:1.3rem; font-weight:700; color:#00923f; margin:0; position:relative; z-index:2;">{int(max_small_orders)} 单</p>
                        </div>
                    </div>
                </div>
                ''', unsafe_allow_html=True)
            else:
                st.markdown('''
                <div class="glass-card" style="padding:1rem; text-align:center;">
                    <h4 style="margin-bottom:0.8rem; color:var(--color-accent-green);">🌟 量采之星</h4>
                    <div style="padding: 1rem;">
                        <p>没有找到小单数据</p>
                    </div>
                </div>
                ''', unsafe_allow_html=True)
        else:
            st.markdown('''
            <div class="glass-card" style="padding:1rem; text-align:center;">
                <h4 style="margin-bottom:0.8rem; color:var(--color-accent-green);">🌟 量采之星</h4>
                <div style="padding: 1rem;">
                    <p>数据中缺少必要的列</p>
                </div>
            </div>
            ''', unsafe_allow_html=True)

    # 3. 总采至尊 - 总采购金额最高者
    with row2_col1:
        # 确定姓名列名
        name_col = '姓名' if '姓名' in stats_df.columns else '报价员'

        # 检查是否有"总采购金额"列
        if '总采购金额' in stats_df.columns and name_col in stats_df.columns:
            # 排除"合计"行
            stats_no_total = stats_df[stats_df[name_col] != '合计'].copy()

            if not stats_no_total.empty:
                # 找出总采购金额最高的人
                max_total_amount = stats_no_total['总采购金额'].max()
                top_total_person = stats_no_total[stats_no_total['总采购金额'] == max_total_amount].iloc[0]

                # 显示总采至尊卡片
                st.markdown(f'''
                <div class="glass-card" style="padding:1.2rem; text-align:center; border-radius:12px; box-shadow:0 10px 25px rgba(191, 90, 242, 0.18); border:1px solid rgba(191, 90, 242, 0.2); position:relative; overflow:hidden; background:linear-gradient(135deg, rgba(191, 90, 242, 0.05) 0%, rgba(255, 255, 255, 0.9) 50%, rgba(191, 90, 242, 0.05) 100%);">
                    <div style="position:absolute; top:0; right:0; width:100%; height:5px; background:linear-gradient(90deg, transparent, rgba(191, 90, 242, 0.5), transparent);"></div>
                    <div style="position:absolute; bottom:0; left:0; width:100%; height:5px; background:linear-gradient(90deg, transparent, rgba(191, 90, 242, 0.5), transparent);"></div>
                    <h4 style="margin-bottom:1rem; color:var(--color-accent-purple); font-size:1.4rem; font-weight:700; text-transform:uppercase; letter-spacing:1px;">👑 总采至尊</h4>
                    <div style="display:flex; flex-direction:column; align-items:center; position:relative; z-index:2;">
                        <div style="width:100px; height:100px; border-radius:20px; background:linear-gradient(135deg, #BF5AF2, #9529ce); display:flex; align-items:center; justify-content:center; margin-bottom:1rem; box-shadow:0 10px 25px rgba(191, 90, 242, 0.3); clip-path:polygon(50% 0%, 100% 25%, 100% 75%, 50% 100%, 0% 75%, 0% 25%);">
                            <span style="font-size:2.4rem; color:white;">💎</span>
                        </div>
                        <h3 style="margin:0.8rem 0; font-size:1.6rem; background:linear-gradient(90deg, #9529ce, #BF5AF2); -webkit-background-clip:text; -webkit-text-fill-color:transparent; font-weight:700;">{top_total_person[name_col]}</h3>
                        <div style="background:rgba(191, 90, 242, 0.12); padding:0.5rem 1.5rem; border-radius:30px; margin-top:0.5rem; border:1px solid rgba(191, 90, 242, 0.2);">
                            <p style="font-size:1.3rem; font-weight:700; color:#9529ce; margin:0;">{max_total_amount / 10000:.2f} 万元</p>
                        </div>
                    </div>
                </div>
                ''', unsafe_allow_html=True)
            else:
                st.markdown('''
                <div class="glass-card" style="padding:1rem; text-align:center;">
                    <h4 style="margin-bottom:0.8rem; color:var(--color-accent-purple);">👑 总采至尊</h4>
                    <div style="padding: 1rem;">
                        <p>没有找到总采购金额数据</p>
                    </div>
                </div>
                ''', unsafe_allow_html=True)
        else:
            st.markdown('''
            <div class="glass-card" style="padding:1rem; text-align:center;">
                <h4 style="margin-bottom:0.8rem; color:var(--color-accent-purple);">👑 总采至尊</h4>
                <div style="padding: 1rem;">
                    <p>数据中缺少必要的列</p>
                </div>
            </div>
            ''', unsafe_allow_html=True)

    # 4. 小单拿货王 - 有效拿货量最高者
    with row2_col2:
        # 检查是否有拿货统计数据和"有效拿货量"列
        if delivery_df is not None and '有效拿货量' in delivery_df.columns:
            # 确定拿货统计的姓名列
            delivery_name_col = '姓名' if '姓名' in delivery_df.columns else '拿货员'

            if delivery_name_col in delivery_df.columns:
                # 排除"合计"行
                delivery_no_total = delivery_df[delivery_df[delivery_name_col] != '合计'].copy()

                if not delivery_no_total.empty:
                    # 找出有效拿货量最高的人
                    max_pickup = delivery_no_total['有效拿货量'].max()
                    top_pickup_person = delivery_no_total[delivery_no_total['有效拿货量'] == max_pickup].iloc[0]

                    # 显示小单拿货王卡片
                    st.markdown(f'''
                    <div class="glass-card" style="padding:1.2rem; text-align:center; border-radius:16px; box-shadow:0 10px 25px rgba(255, 214, 10, 0.18); border:1px solid rgba(255, 214, 10, 0.2); position:relative; overflow:hidden;">
                        <div style="position:absolute; top:0; left:0; width:100%; height:100%; background:repeating-linear-gradient(45deg, rgba(255, 214, 10, 0.03), rgba(255, 214, 10, 0.03) 10px, rgba(255, 214, 10, 0) 10px, rgba(255, 214, 10, 0) 20px);"></div>
                        <h4 style="margin-bottom:1rem; color:#e0a800; font-size:1.4rem; font-weight:700; position:relative;">
                            <span style="display:inline-block; animation:shake 3s infinite;">🚚</span> 小单拿货王
                        </h4>
                        <div style="display:flex; flex-direction:column; align-items:center; position:relative; z-index:2;">
                            <div style="width:90px; height:90px; border-radius:50% 50% 5px 50%; transform:rotate(-5deg); background:linear-gradient(135deg, #FFD60A, #e0a800); display:flex; align-items:center; justify-content:center; margin-bottom:1rem; box-shadow:0 8px 20px rgba(255, 214, 10, 0.3);">
                                <span style="font-size:2.2rem; color:white;">🔥</span>
                            </div>
                            <h3 style="margin:0.8rem 0; font-size:1.6rem; color:#e0a800;">{top_pickup_person[delivery_name_col]}</h3>
                            <div style="background:rgba(255, 214, 10, 0.12); padding:0.5rem 1rem; border-radius:30px; margin-top:0.5rem; position:relative;">
                                <div style="position:absolute; right:-5px; top:-5px; width:20px; height:20px; background:#FFD60A; opacity:0.2; border-radius:50%;"></div>
                                <p style="font-size:1.3rem; font-weight:700; color:#e0a800; margin:0;">{float(max_pickup):.1f} 单</p>
                            </div>
                        </div>
                    </div>
                    ''', unsafe_allow_html=True)
                else:
                    st.markdown('''
                    <div class="glass-card" style="padding:1rem; text-align:center;">
                        <h4 style="margin-bottom:0.8rem; color:var(--color-accent-yellow);">🚚 小单拿货王</h4>
                        <div style="padding: 1rem;">
                            <p>没有找到拿货数据</p>
                        </div>
                    </div>
                    ''', unsafe_allow_html=True)
            else:
                st.markdown('''
                <div class="glass-card" style="padding:1rem; text-align:center;">
                    <h4 style="margin-bottom:0.8rem; color:var(--color-accent-yellow);">🚚 小单拿货王</h4>
                    <div style="padding: 1rem;">
                        <p>拿货数据中缺少姓名列</p>
                    </div>
                </div>
                ''', unsafe_allow_html=True)
        else:
            st.markdown('''
            <div class="glass-card" style="padding:1rem; text-align:center;">
                <h4 style="margin-bottom:0.8rem; color:var(--color-accent-yellow);">🚚 小单拿货王</h4>
                <div style="padding: 1rem;">
                    <p>未找到拿货统计数据或数据中缺少必要的列</p>
                </div>
            </div>
            ''', unsafe_allow_html=True)

    # 关闭外部容器
    st.markdown('</div>', unsafe_allow_html=True)


# 显示采购详情页面
def show_purchase_detail_page():
    # 检查是否已加载数据
    if not st.session_state.get('data_loaded', False) or st.session_state['excel_data'] is None:
        show_warning_card("请先上传采购统计表数据")
        return

    # 获取采购详情数据
    excel_data = st.session_state['excel_data']
    details_df = excel_data.get('采购详情')

    if details_df is None:
        show_error_card("未能找到采购详情数据")
        return

    # 创建选项卡
    tab1, tab2 = st.tabs(["大单详情", "小单详情"])

    # 提取姓名列(可能是"姓名"或"报价员"列)
    name_col = '姓名' if '姓名' in details_df.columns else ('报价员' if '报价员' in details_df.columns else None)

    # 大单详情选项卡
    with tab1:
        # 创建包含标题的卡片
        st.markdown('''
        <div class="glass-card compact-card">
            <h3 class="card-title">大单采购明细</h3>
        ''', unsafe_allow_html=True)

        # 提取大单数据
        large_order_columns = [
            name_col, '大单序号', '大单开始日期', '大单结束日期',
            '大单采购单号', '大单产品编码', '大单采购金额'
        ]

        # 确保所有必要的列都存在
        missing_cols = [col for col in large_order_columns if col not in details_df.columns]
        if missing_cols or name_col is None:
            st.error(f"数据中缺少必要的列: {', '.join(missing_cols if missing_cols else ['姓名/报价员'])}")
        else:
            # 筛选有大单序号的行
            large_orders = details_df[details_df['大单序号'].notna() & (details_df['大单序号'] != '')]

            if large_orders.empty:
                st.info("没有大单数据")
            else:
                # 只保留需要的列
                large_orders = large_orders[large_order_columns]

                # 添加姓名筛选功能
                all_names = sorted(large_orders[name_col].unique())
                all_names.insert(0, "全部")  # 添加"全部"选项

                selected_name = st.selectbox("选择报价员筛选大单", all_names, key="large_order_name")

                # 根据选择筛选数据
                if selected_name != "全部":
                    filtered_large_orders = large_orders[large_orders[name_col] == selected_name]
                else:
                    filtered_large_orders = large_orders

                # 显示筛选后的数据
                if not filtered_large_orders.empty:
                    st.dataframe(filtered_large_orders, use_container_width=True, height=450, hide_index=True)
                else:
                    st.info(f"没有找到{selected_name}的大单数据")

        # 关闭卡片
        st.markdown('</div>', unsafe_allow_html=True)

    # 小单详情选项卡
    with tab2:
        # 创建包含标题的卡片
        st.markdown('''
        <div class="glass-card compact-card">
            <h3 class="card-title">小单采购明细</h3>
        ''', unsafe_allow_html=True)

        # 提取小单数据
        small_order_columns = [
            name_col, '小单序号', '小单开始日期', '小单结束日期',
            '小单采购单号', '小单产品编码', '小单采购金额'
        ]

        # 确保所有必要的列都存在
        missing_cols = [col for col in small_order_columns if col not in details_df.columns]
        if missing_cols or name_col is None:
            st.error(f"数据中缺少必要的列: {', '.join(missing_cols if missing_cols else ['姓名/报价员'])}")
        else:
            # 筛选有小单序号的行
            small_orders = details_df[details_df['小单序号'].notna() & (details_df['小单序号'] != '')]

            if small_orders.empty:
                st.info("没有小单数据")
            else:
                # 只保留需要的列
                small_orders = small_orders[small_order_columns]

                # 添加姓名筛选功能
                all_names = sorted(small_orders[name_col].unique())
                all_names.insert(0, "全部")  # 添加"全部"选项

                selected_name = st.selectbox("选择报价员筛选小单", all_names, key="small_order_name")

                # 根据选择筛选数据
                if selected_name != "全部":
                    filtered_small_orders = small_orders[small_orders[name_col] == selected_name]
                else:
                    filtered_small_orders = small_orders

                # 显示筛选后的数据
                if not filtered_small_orders.empty:
                    st.dataframe(filtered_small_orders, use_container_width=True, height=450, hide_index=True)
                else:
                    st.info(f"没有找到{selected_name}的小单数据")

        # 关闭卡片
    st.markdown('</div>', unsafe_allow_html=True)


# 显示采购统计页面
def show_purchase_stats_page():
    # 检查是否已加载数据
    if not st.session_state.get('data_loaded', False) or st.session_state['excel_data'] is None:
        show_warning_card("请先上传采购统计表数据")
        return

    # 获取采购统计数据
    excel_data = st.session_state['excel_data']
    stats_df = excel_data.get('采购统计')

    if stats_df is None:
        show_error_card("未能找到采购统计数据")
        return

    # 第一部分：本月采购汇总标题
    st.markdown("""
    <h2 style="margin: 20px 0 15px 0; padding: 0; font-size: 28px; font-weight: 700; color: #1D1D1F; font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;">
        一、📊 本月采购汇总
    </h2>
    """, unsafe_allow_html=True)

    # 直接显示数据，不使用额外卡片
    # 获取列名
    name_col = '姓名' if '姓名' in stats_df.columns else ('报价员' if '报价员' in stats_df.columns else None)

    if name_col and '合计' in stats_df[name_col].values:
        # 提取合计行数据
        total_row = stats_df[stats_df[name_col] == '合计'].copy()

        # 创建4个指标的列容器
        cols = st.columns(4)

        # 1. 本月大单数 - 直接使用确定的列名
        if '大单数' in total_row.columns:
            value = total_row['大单数'].values[0]
            # 转换为字符串，然后尝试转换为数字
            value_str = str(value)
            try:
                if value_str and value_str.strip() and value_str.replace('.', '', 1).isdigit():
                    value_num = float(value_str)
                    cols[0].metric("本月大单数", f"{int(value_num)}")
                else:
                    cols[0].metric("本月大单数", value_str)
            except:
                cols[0].metric("本月大单数", value_str)
        else:
            cols[0].metric("本月大单数", "列不存在")

        # 2. 本月小单数 - 直接使用确定的列名
        if '小单数' in total_row.columns:
            value = total_row['小单数'].values[0]
            # 转换为字符串，然后尝试转换为数字
            value_str = str(value)
            try:
                if value_str and value_str.strip() and value_str.replace('.', '', 1).isdigit():
                    value_num = float(value_str)
                    cols[1].metric("本月小单数", f"{int(value_num)}")
                else:
                    cols[1].metric("本月小单数", value_str)
            except:
                cols[1].metric("本月小单数", value_str)
        else:
            cols[1].metric("本月小单数", "列不存在")

        # 3. 本月总采购额 - 已经正确显示，保持不变
        if '总采购金额' in total_row.columns:
            value = total_row['总采购金额'].values[0]
            try:
                value_num = float(str(value).replace(',', ''))
                cols[2].metric("本月总采购额", f"¥{value_num:,.2f}")
            except:
                cols[2].metric("本月总采购额", f"{value}")
        else:
            cols[2].metric("本月总采购额", "列不存在")

        # 4. 本月目标完成进度 - 已经正确显示，保持不变
        if '采购业绩完成进度' in total_row.columns:
            value = total_row['采购业绩完成进度'].values[0]
            try:
                # 处理百分比字符串情况
                value_str = str(value)
                if '%' in value_str:
                    value_num = float(value_str.replace('%', '')) / 100
                else:
                    value_num = float(value_str)

                # 根据完成进度设置不同的颜色
                if value_num < 0.66:
                    progress_color = "red"
                elif value_num < 1.0:
                    progress_color = "orange"
                else:
                    progress_color = "green"

                # 显示带颜色的进度
                cols[3].markdown(
                    f"""
                    <div style="text-align: center;">
                        <p style="font-size: 14px; color: gray; margin-bottom: 0;">本月目标完成进度</p>
                        <p style="font-size: 24px; font-weight: bold; color: {progress_color}; margin: 0;">{value_num:.1%}</p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
            except:
                cols[3].metric("本月目标完成进度", f"{value}")
        else:
            cols[3].metric("本月目标完成进度", "列不存在")
    else:
        st.warning("未找到合计行数据，无法显示本月采购汇总")

    # 第二部分：员工采购额排名标题
    st.markdown("""
    <h2 style="margin: 30px 0 15px 0; padding: 0; font-size: 28px; font-weight: 700; color: #1D1D1F; font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;">
        二、🏆 员工采购额排名
    </h2>
    """, unsafe_allow_html=True)

    # 排名数据处理和展示
    name_col = '姓名' if '姓名' in stats_df.columns else ('报价员' if '报价员' in stats_df.columns else None)

    if name_col and '总采购金额' in stats_df.columns:
        # 去除合计行
        employee_df = stats_df[stats_df[name_col] != '合计'].copy()

        # 确保总采购金额是数值型
        employee_df['总采购金额'] = pd.to_numeric(employee_df['总采购金额'], errors='coerce')

        # 按总采购金额降序排序
        employee_df = employee_df.sort_values(by='总采购金额', ascending=False).reset_index(drop=True)

        # 计算员工数量
        employee_count = len(employee_df)

        # 使用自定义排名表格而不是matplotlib图表，避免中文和emoji显示问题
        st.markdown("""
        <style>
        .custom-rank-table {
            width: 100%;
            border-collapse: collapse;
            font-family: "微软雅黑", "Microsoft YaHei", sans-serif;
        }
        .custom-rank-row {
            display: flex;
            align-items: center;
            margin-bottom: 8px;
            padding: 10px;
            background-color: rgba(255, 255, 255, 0.8);
            border-radius: 10px;
            box-shadow: 0 2px 5px rgba(0, 0, 0, 0.05);
        }
        .custom-rank-row:hover {
            background-color: rgba(255, 255, 255, 0.95);
            transform: translateY(-2px);
            transition: all 0.2s ease;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.08);
        }
        .rank-number {
            width: 32px;
            height: 32px;
            border-radius: 50%;
            color: white;
            font-weight: bold;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-right: 10px;
            flex-shrink: 0;
        }
        .rank-top-1 {
            background: linear-gradient(135deg, #FFD700, #FFA500);
        }
        .rank-top-2 {
            background: linear-gradient(135deg, #C0C0C0, #A0A0A0);
        }
        .rank-top-3 {
            background: linear-gradient(135deg, #CD7F32, #B87333);
        }
        .rank-normal {
            background: linear-gradient(135deg, #0A84FF, #0055c1);
        }
        .rank-bottom {
            background: linear-gradient(135deg, #FF453A, #c13000);
        }
        .rank-name {
            font-weight: 500;
            width: 80px;
            flex-shrink: 0;
            font-size: 16px;
        }
        .rank-bar-container {
            flex-grow: 1;
            height: 24px;
            background: rgba(0, 0, 0, 0.05);
            border-radius: 12px;
            margin: 0 10px;
            position: relative;
            overflow: hidden;
        }
        .rank-bar {
            height: 100%;
            border-radius: 12px;
            position: absolute;
            top: 0;
            left: 0;
        }
        .bar-top-1 {
            background: linear-gradient(90deg, #FFD700, #FFA500);
        }
        .bar-top-2 {
            background: linear-gradient(90deg, #C0C0C0, #A0A0A0);
        }
        .bar-top-3 {
            background: linear-gradient(90deg, #CD7F32, #B87333);
        }
        .bar-normal {
            background: linear-gradient(90deg, #0A84FF, #0055c1);
        }
        .bar-bottom {
            background: linear-gradient(90deg, #FF453A, #c13000);
        }
        .rank-amount {
            width: 100px;
            text-align: right;
            font-weight: bold;
            flex-shrink: 0;
            font-size: 15px;
        }
        .rank-icon {
            margin-left: 10px;
            font-size: 22px;
            width: 30px;
        }
        </style>
        """, unsafe_allow_html=True)

        # 计算最大金额，用于条形图百分比
        max_amount = employee_df['总采购金额'].max()

        # 遍历员工创建各个排名项
        for i, row in employee_df.iterrows():
            rank = i + 1
            name = row[name_col]
            amount = row['总采购金额']
            percentage = (amount / max_amount) * 100 if max_amount > 0 else 0

            # 根据排名确定样式和图标
            if rank == 1:
                rank_class = "rank-top-1"
                bar_class = "bar-top-1"
                icon_html = '<span class="rank-icon">🏆</span>'
            elif rank == 2:
                rank_class = "rank-top-2"
                bar_class = "bar-top-2"
                icon_html = '<span class="rank-icon">🥈</span>'
            elif rank == 3:
                rank_class = "rank-top-3"
                bar_class = "bar-top-3"
                icon_html = '<span class="rank-icon">🥉</span>'
            # 修改逻辑：只有当员工数量大于等于7人时，才对最后三名显示警告标记
            elif employee_count >= 7 and employee_count - rank < 3:  # 最后三名（仅当员工数量大于等于7人时）
                rank_class = "rank-bottom"
                bar_class = "bar-bottom"
                icon_html = '<span class="rank-icon">⚠️</span>'
            else:
                rank_class = "rank-normal"
                bar_class = "bar-normal"
                icon_html = '<span class="rank-icon"></span>'  # 空白占位符，保持布局一致

            # 单独渲染每一行
            st.markdown(f"""
            <div class="custom-rank-row">
                <div class="rank-number {rank_class}">{rank}</div>
                <div class="rank-name">{name}</div>
                <div class="rank-bar-container">
                    <div class="rank-bar {bar_class}" style="width: {percentage}%;"></div>
                </div>
                <div class="rank-amount">¥{amount:,.2f}</div>
                {icon_html}
            </div>
            """, unsafe_allow_html=True)
    else:
        st.warning("无法显示员工采购额排名：缺少必要的列或数据")

    # 第三部分：周采购数据分析标题
    st.markdown("""
    <h2 style="margin: 30px 0 15px 0; padding: 0; font-size: 28px; font-weight: 700; color: #1D1D1F; font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;">
        三、📈 周采购数据分析
    </h2>
    """, unsafe_allow_html=True)

    # 检查是否有周度数据列
    if stats_df is not None and '合计' in stats_df[name_col].values:
        # 从合计行提取数据
        total_row = stats_df[stats_df[name_col] == '合计'].copy().reset_index(drop=True)

        # 初始化周度数据列表
        week_cols_amount = [col for col in total_row.columns if '周采购金额' in col]
        week_cols_orders = [col for col in total_row.columns if '周小单数' in col]

        # 提取周度数据
        if week_cols_amount and week_cols_orders:
            # 创建周次标签
            weeks = [f"第{i + 1}周" for i in range(len(week_cols_amount))]

            # 提取周采购金额数据
            week_amounts = []
            for col in week_cols_amount:
                try:
                    value = total_row[col].values[0]
                    if pd.notna(value):
                        # 处理可能的字符串格式
                        if isinstance(value, str):
                            value = float(value.replace(',', '').replace('¥', ''))
                        week_amounts.append(value)
                    else:
                        week_amounts.append(0)
                except:
                    week_amounts.append(0)

            # 提取周小单数数据
            week_orders = []
            for col in week_cols_orders:
                try:
                    value = total_row[col].values[0]
                    if pd.notna(value):
                        if isinstance(value, str) and value.strip() and value.replace('.', '', 1).isdigit():
                            value = float(value)
                        week_orders.append(value)
                    else:
                        week_orders.append(0)
                except:
                    week_orders.append(0)

            # 创建DataFrame用于表格显示 - 周采购金额
            week_amounts_df = pd.DataFrame({
                '周次': weeks,
                '采购金额': [f"¥{amount:,.2f}" for amount in week_amounts],
            })

            # 创建DataFrame用于表格显示 - 周小单数
            week_orders_df = pd.DataFrame({
                '周次': weeks,
                '小单数': week_orders
            })

            # 使用Plotly创建交互式图表
            import plotly.graph_objects as go

            # 1. 周采购金额折线图
            fig_amount = go.Figure()

            # 添加金额折线图
            fig_amount.add_trace(
                go.Scatter(
                    x=weeks,
                    y=week_amounts,
                    name="周采购金额",
                    line=dict(color="#0A84FF", width=3),
                    mode='lines+markers',
                    marker=dict(size=10, symbol='circle', color="#0A84FF",
                                line=dict(color='white', width=2))
                )
            )

            # 添加图表样式 - 只使用有效的Plotly属性
            fig_amount.update_layout(
                title_text="周采购金额趋势分析<br><span style='font-size:12px;color:#86868B'>点击员工图例，隐藏对应数据线</span>",
                font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif",
                          size=14),
                plot_bgcolor='rgba(255, 255, 255, 0.8)',  # 设置绘图区背景为白色
                paper_bgcolor='white',  # 设置图表背景为白色
                hovermode="x unified",
                margin=dict(l=30, r=30, t=100, b=30),  # 增加顶部边距
                # 添加边框效果
                shapes=[
                    # 外框
                    dict(
                        type='rect',
                        xref='paper', yref='paper',
                        x0=-0.05, y0=-0.05, x1=1.05, y1=1.05,
                        line=dict(color='rgba(0, 0, 0, 0.08)', width=1),
                        fillcolor='rgba(0, 0, 0, 0)',
                        layer='below'
                    )
                ],
                # 设置图表大小
                width=None,  # 自适应宽度
                height=650,  # 增加高度
                # 添加图表标题样式
                title={
                    'y': 0.95,
                    'x': 0.5,
                    'xanchor': 'center',
                    'yanchor': 'top',
                    'font': dict(size=18, color="#1D1D1F")
                },
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.12,  # 增加与图表的距离
                    xanchor="center",
                    x=0.5
                )
            )

            # 更新轴标签和样式
            fig_amount.update_xaxes(
                title_text="周次",
                gridcolor='rgba(0, 0, 0, 0.05)',
                showline=False,  # 取消轴边框线
                linewidth=1,
                linecolor='rgba(0, 0, 0, 0.1)',
                tickfont=dict(family="SF Pro Display", size=12)
            )

            # 更新Y轴 (金额)
            fig_amount.update_yaxes(
                title_text="采购金额 (元)",
                gridcolor='rgba(0, 0, 0, 0.05)',
                showline=False,  # 取消轴边框线
                linewidth=1,
                linecolor='rgba(0, 0, 0, 0.1)',
                tickformat=",.0f"
            )

            # 显示图表
            st.plotly_chart(fig_amount, use_container_width=True)

            # 显示金额表格数据
            st.markdown("<h4 style='text-align: center; margin: 10px 0;'>周度采购金额明细</h4>",
                        unsafe_allow_html=True)
            st.dataframe(week_amounts_df, use_container_width=True, hide_index=True)

            # 2. 周小单数折线图
            fig_orders = go.Figure()

            # 添加小单数折线图
            fig_orders.add_trace(
                go.Scatter(
                    x=weeks,
                    y=week_orders,
                    name="周小单数",
                    line=dict(color="#30D158", width=3),
                    mode='lines+markers',
                    marker=dict(size=10, symbol='diamond', color="#30D158",
                                line=dict(color='white', width=2))
                )
            )

            # 添加图表样式 - 只使用有效的Plotly属性
            fig_orders.update_layout(
                title_text="周小单数趋势分析<br><span style='font-size:12px;color:#86868B'>点击员工图例，隐藏对应数据线</span>",
                font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif",
                          size=14),
                plot_bgcolor='rgba(255, 255, 255, 0.8)',  # 设置绘图区背景为白色
                paper_bgcolor='white',  # 设置图表背景为白色
                hovermode="x unified",
                margin=dict(l=30, r=30, t=100, b=30),  # 增加顶部边距
                # 添加边框效果
                shapes=[
                    # 外框
                    dict(
                        type='rect',
                        xref='paper', yref='paper',
                        x0=-0.05, y0=-0.05, x1=1.05, y1=1.05,
                        line=dict(color='rgba(0, 0, 0, 0.08)', width=1),
                        fillcolor='rgba(0, 0, 0, 0)',
                        layer='below'
                    )
                ],
                # 设置图表大小
                width=None,  # 自适应宽度
                height=650,  # 增加高度
                # 添加图表标题样式
                title={
                    'y': 0.95,
                    'x': 0.5,
                    'xanchor': 'center',
                    'yanchor': 'top',
                    'font': dict(size=18, color="#1D1D1F")
                },
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.12,  # 增加与图表的距离
                    xanchor="center",
                    x=0.5
                )
            )

            # 更新轴标签和样式
            fig_orders.update_xaxes(
                title_text="周次",
                gridcolor='rgba(0, 0, 0, 0.05)',
                showline=False,  # 取消轴边框线
                linewidth=1,
                linecolor='rgba(0, 0, 0, 0.1)',
                tickfont=dict(family="SF Pro Display", size=12)
            )

            # 更新Y轴 (订单数)
            fig_orders.update_yaxes(
                title_text="小单数量",
                gridcolor='rgba(0, 0, 0, 0.05)',
                showline=False,  # 取消轴边框线
                linewidth=1,
                linecolor='rgba(0, 0, 0, 0.1)',
                tickformat=",d"
            )

            # 显示图表
            st.plotly_chart(fig_orders, use_container_width=True)

            # 显示订单表格数据
            st.markdown("<h4 style='text-align: center; margin: 10px 0;'>周度小单数明细</h4>",
                        unsafe_allow_html=True)
            st.dataframe(week_orders_df, use_container_width=True, hide_index=True)

        else:
            st.warning("未找到周度采购数据列。请确保表格中包含'第X周采购金额'和'第X周小单数'列。")
    else:
        st.warning("无法获取采购统计合计行数据")

    # 第四部分：员工周采购数据分析
    st.markdown("""
    <h2 style="margin: 30px 0 15px 0; padding: 0; font-size: 28px; font-weight: 700; color: #1D1D1F; font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;">
        四、👨‍💼 员工周采购数据分析
    </h2>
    """, unsafe_allow_html=True)

    # 检查是否有周度数据列
    if stats_df is not None and name_col is not None:
        # 获取所有员工数据（排除"合计"行）
        employees_df = stats_df[stats_df[name_col] != '合计'].copy()

        if not employees_df.empty:
            # 找出所有周采购金额和周小单数的列
            week_amount_cols = [col for col in employees_df.columns if '周采购金额' in col]
            week_order_cols = [col for col in employees_df.columns if '周小单数' in col]

            if week_amount_cols and week_order_cols:
                # 创建周次标签
                weeks = [f"第{i + 1}周" for i in range(len(week_amount_cols))]

                # 1. 员工周采购金额折线图
                fig_emp_amounts = go.Figure()

                # 为每个员工添加一条折线
                for i, row in employees_df.iterrows():
                    employee_name = row[name_col]
                    # 提取该员工的周采购金额数据
                    emp_amounts = []
                    for col in week_amount_cols:
                        try:
                            value = row[col]
                            if pd.notna(value):
                                # 处理可能的字符串格式
                                if isinstance(value, str):
                                    value = float(value.replace(',', '').replace('¥', ''))
                                emp_amounts.append(value)
                            else:
                                emp_amounts.append(0)
                        except:
                            emp_amounts.append(0)

                    # 添加该员工的折线
                    fig_emp_amounts.add_trace(
                        go.Scatter(
                            x=weeks,
                            y=emp_amounts,
                            name=employee_name,
                            mode='lines+markers',
                            marker=dict(size=8)
                        )
                    )

                # 添加图表样式
                fig_emp_amounts.update_layout(
                    title_text="员工周采购金额趋势对比<br><span style='font-size:12px;color:#86868B'>点击员工图例，隐藏对应数据线</span>",
                    font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif", size=14),
                    plot_bgcolor='rgba(255, 255, 255, 0.8)',
                    paper_bgcolor='white',
                    hovermode="x unified",
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",  # 确保图例从图表顶部向上延伸
                        y=1.05,  # 将图例放在图表外部，与标题保持相等距离
                        xanchor="center",
                        x=0.485  # 从0.48调整到0.485，使图例与标题更精确对齐
                    ),
                    margin=dict(l=30, r=30, t=130, b=30),  # 增加顶部边距到130px
                    shapes=[
                        dict(
                            type='rect',
                            xref='paper', yref='paper',
                            x0=-0.05, y0=-0.05, x1=1.05, y1=1.05,
                            line=dict(color='rgba(0, 0, 0, 0.08)', width=1),
                            fillcolor='rgba(0, 0, 0, 0)',
                            layer='below'
                        )
                    ],
                    width=None,
                    height=750,  # 保持图表高度为750px
                    title={
                        'y': 0.95,  # 标题位置保持在图表顶部区域的95%位置
                        'x': 0.5,
                        'xanchor': 'center',
                        'yanchor': 'top',
                        'font': dict(size=18, color="#1D1D1F")
                    }
                )

                # 更新轴标签和样式
                fig_emp_amounts.update_xaxes(
                    title_text="周次",
                    gridcolor='rgba(0, 0, 0, 0.05)',
                    showline=False,
                    linewidth=1,
                    linecolor='rgba(0, 0, 0, 0.1)',
                    tickfont=dict(family="SF Pro Display", size=12)
                )

                fig_emp_amounts.update_yaxes(
                    title_text="采购金额 (元)",
                    gridcolor='rgba(0, 0, 0, 0.05)',
                    showline=False,
                    linewidth=1,
                    linecolor='rgba(0, 0, 0, 0.1)',
                    tickformat=",.0f"
                )

                # 显示图表
                st.plotly_chart(fig_emp_amounts, use_container_width=True)

                # 创建并显示员工周采购金额详细表格
                emp_amounts_data = []
                for i, row in employees_df.iterrows():
                    emp_row = {name_col: row[name_col]}
                    for j, col in enumerate(week_amount_cols):
                        try:
                            value = row[col]
                            if pd.notna(value):
                                if isinstance(value, str):
                                    value = float(value.replace(',', '').replace('¥', ''))
                                emp_row[weeks[j]] = f"¥{value:,.2f}"
                            else:
                                emp_row[weeks[j]] = "¥0.00"
                        except:
                            emp_row[weeks[j]] = "¥0.00"
                    emp_amounts_data.append(emp_row)

                emp_amounts_df = pd.DataFrame(emp_amounts_data)
                st.markdown("<h4 style='text-align: center; margin: 10px 0;'>员工周采购金额明细</h4>",
                            unsafe_allow_html=True)
                st.dataframe(emp_amounts_df, use_container_width=True, hide_index=True)

                # 2. 员工周小单数折线图
                fig_emp_orders = go.Figure()

                # 为每个员工添加一条折线
                for i, row in employees_df.iterrows():
                    employee_name = row[name_col]
                    # 提取该员工的周小单数数据
                    emp_orders = []
                    for col in week_order_cols:
                        try:
                            value = row[col]
                            if pd.notna(value):
                                if isinstance(value, str) and value.strip() and value.replace('.', '', 1).isdigit():
                                    value = float(value)
                                emp_orders.append(value)
                            else:
                                emp_orders.append(0)
                        except:
                            emp_orders.append(0)

                    # 添加该员工的折线
                    fig_emp_orders.add_trace(
                        go.Scatter(
                            x=weeks,
                            y=emp_orders,
                            name=employee_name,
                            mode='lines+markers',
                            marker=dict(size=8)
                        )
                    )

                # 添加图表样式
                fig_emp_orders.update_layout(
                    title_text="员工周小单数趋势对比<br><span style='font-size:12px;color:#86868B'>点击员工图例，隐藏对应数据线</span>",
                    font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif", size=14),
                    plot_bgcolor='rgba(255, 255, 255, 0.8)',
                    paper_bgcolor='white',
                    hovermode="x unified",
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",  # 确保图例从图表顶部向上延伸
                        y=1.05,  # 将图例放在图表外部，与标题保持相等距离
                        xanchor="center",
                        x=0.485  # 从0.48调整到0.485，使图例与标题更精确对齐
                    ),
                    margin=dict(l=30, r=30, t=130, b=30),  # 增加顶部边距到130px
                    shapes=[
                        dict(
                            type='rect',
                            xref='paper', yref='paper',
                            x0=-0.05, y0=-0.05, x1=1.05, y1=1.05,
                            line=dict(color='rgba(0, 0, 0, 0.08)', width=1),
                            fillcolor='rgba(0, 0, 0, 0)',
                            layer='below'
                        )
                    ],
                    width=None,
                    height=750,  # 保持图表高度为750px
                    title={
                        'y': 0.95,  # 标题位置保持在图表顶部区域的95%位置
                        'x': 0.5,
                        'xanchor': 'center',
                        'yanchor': 'top',
                        'font': dict(size=18, color="#1D1D1F")
                    }
                )

                # 更新轴标签和样式
                fig_emp_orders.update_xaxes(
                    title_text="周次",
                    gridcolor='rgba(0, 0, 0, 0.05)',
                    showline=False,
                    linewidth=1,
                    linecolor='rgba(0, 0, 0, 0.1)',
                    tickfont=dict(family="SF Pro Display", size=12)
                )

                fig_emp_orders.update_yaxes(
                    title_text="小单数量",
                    gridcolor='rgba(0, 0, 0, 0.05)',
                    showline=False,
                    linewidth=1,
                    linecolor='rgba(0, 0, 0, 0.1)',
                    tickformat=",d"
                )

                # 显示图表
                st.plotly_chart(fig_emp_orders, use_container_width=True)

                # 创建并显示员工周小单数详细表格
                emp_orders_data = []
                for i, row in employees_df.iterrows():
                    emp_row = {name_col: row[name_col]}
                    for j, col in enumerate(week_order_cols):
                        try:
                            value = row[col]
                            if pd.notna(value):
                                if isinstance(value, str) and value.strip() and value.replace('.', '', 1).isdigit():
                                    value = float(value)
                                emp_row[weeks[j]] = int(value)
                            else:
                                emp_row[weeks[j]] = 0
                        except:
                            emp_row[weeks[j]] = 0
                    emp_orders_data.append(emp_row)

                emp_orders_df = pd.DataFrame(emp_orders_data)
                st.markdown("<h4 style='text-align: center; margin: 10px 0;'>员工周小单数明细</h4>",
                            unsafe_allow_html=True)
                st.dataframe(emp_orders_df, use_container_width=True, hide_index=True)
            else:
                st.warning("未找到周度采购数据列。请确保表格中包含'第X周采购金额'和'第X周小单数'列。")
        else:
            st.warning("没有找到员工数据。")
    else:
        st.warning("无法获取员工数据进行分析")

    # 第五部分：员工大小单占比
    st.markdown("""
    <h2 style="margin: 30px 0 15px 0; padding: 0; font-size: 28px; font-weight: 700; color: #1D1D1F; font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;">
        五、📊 员工大小单占比
    </h2>
    """, unsafe_allow_html=True)

    # 检查是否有员工数据
    if stats_df is not None and name_col is not None:
        # 获取所有员工数据（排除"合计"行）
        employees_df = stats_df[stats_df[name_col] != '合计'].copy()

        # 检查是否有大单数和小单数列
        if '大单数' in employees_df.columns and '小单数' in employees_df.columns:
            # 确保大单数和小单数是数值型
            employees_df['大单数'] = pd.to_numeric(employees_df['大单数'], errors='coerce').fillna(0)
            employees_df['小单数'] = pd.to_numeric(employees_df['小单数'], errors='coerce').fillna(0)

            # 排序员工，按总单数（大单数+小单数）降序排列
            employees_df['总单数'] = employees_df['大单数'] + employees_df['小单数']
            employees_df = employees_df.sort_values(by='总单数', ascending=False).reset_index(drop=True)

            # 提取员工姓名、大单数和小单数
            employee_names = employees_df[name_col].tolist()
            large_orders = employees_df['大单数'].tolist()
            small_orders = employees_df['小单数'].tolist()

            # 创建分组柱状图
            import plotly.graph_objects as go

            # 创建图表
            fig_orders = go.Figure()

            # 添加大单数柱状图
            fig_orders.add_trace(go.Bar(
                x=employee_names,
                y=large_orders,
                name='大单数',
                marker_color='#0A84FF',
                hovertemplate='%{x}: %{y}个大单<extra></extra>'
            ))

            # 添加小单数柱状图
            fig_orders.add_trace(go.Bar(
                x=employee_names,
                y=small_orders,
                name='小单数',
                marker_color='#30D158',
                hovertemplate='%{x}: %{y}个小单<extra></extra>'
            ))

            # 更新布局
            fig_orders.update_layout(
                title_text='员工大小单数量分布<br><span style="font-size:12px;color:#86868B">鼠标悬停查看详细数据</span>',
                barmode='group',
                font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif", size=14),
                plot_bgcolor='rgba(255, 255, 255, 0.8)',
                paper_bgcolor='white',
                hovermode="x unified",
                margin=dict(l=30, r=30, t=100, b=100),
                shapes=[
                    dict(
                        type='rect',
                        xref='paper', yref='paper',
                        x0=-0.05, y0=-0.05, x1=1.05, y1=1.05,
                        line=dict(color='rgba(0, 0, 0, 0.08)', width=1),
                        fillcolor='rgba(0, 0, 0, 0)',
                        layer='below'
                    )
                ],
                width=None,
                height=600,
                title={
                    'y': 0.95,
                    'x': 0.5,
                    'xanchor': 'center',
                    'yanchor': 'top',
                    'font': dict(size=18, color="#1D1D1F")
                },
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.05,
                    xanchor="center",
                    x=0.485  # 从0.48调整到0.485，使图例与标题更精确对齐
                ),
                xaxis_tickangle=-45
            )

            # 更新x轴
            fig_orders.update_xaxes(
                title_text="员工",
                gridcolor='rgba(0, 0, 0, 0.05)',
                showline=False,
                linewidth=1,
                linecolor='rgba(0, 0, 0, 0.1)',
                tickfont=dict(family="SF Pro Display", size=12)
            )

            # 更新y轴
            fig_orders.update_yaxes(
                title_text="单数",
                gridcolor='rgba(0, 0, 0, 0.05)',
                showline=False,
                linewidth=1,
                linecolor='rgba(0, 0, 0, 0.1)',
                tickformat=",d"
            )

            # 显示图表
            st.plotly_chart(fig_orders, use_container_width=True)

            # 计算占比数据
            employees_df['大单占比'] = employees_df['大单数'] / employees_df['总单数'] * 100
            employees_df['小单占比'] = employees_df['小单数'] / employees_df['总单数'] * 100

            # 创建数据表格
            table_data = employees_df[[name_col, '大单数', '小单数', '总单数', '大单占比', '小单占比']].copy()
            # 格式化占比为百分比
            table_data['大单占比'] = table_data['大单占比'].apply(lambda x: f"{x:.1f}%")
            table_data['小单占比'] = table_data['小单占比'].apply(lambda x: f"{x:.1f}%")

            # 重命名列
            table_data.columns = ['员工', '大单数', '小单数', '总单数', '大单占比', '小单占比']

            # 显示表格
            st.markdown("<h4 style='text-align: center; margin: 10px 0;'>员工大小单数据明细</h4>",
                        unsafe_allow_html=True)
            st.dataframe(table_data, use_container_width=True, hide_index=True)

            # 计算全员平均值
            avg_large = employees_df['大单数'].mean()
            avg_small = employees_df['小单数'].mean()
            total_large = employees_df['大单数'].sum()
            total_small = employees_df['小单数'].sum()
            total_ratio = total_large / (total_large + total_small) * 100

            # 添加平均值和总体占比信息
            col1, col2, col3 = st.columns(3)

            with col1:
                st.metric("平均大单数", f"{avg_large:.1f}")

            with col2:
                st.metric("平均小单数", f"{avg_small:.1f}")

            with col3:
                st.metric("全员大单占比", f"{total_ratio:.1f}%")

        else:
            st.warning("数据中缺少'大单数'或'小单数'列，无法创建大小单占比分析。")
    else:
        st.warning("无法获取员工数据进行分析")

    # 第六部分：员工详情
    st.markdown("""
    <h2 style="margin: 30px 0 15px 0; padding: 0; font-size: 28px; font-weight: 700; color: #1D1D1F; font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;">
        六、👨‍💼 员工详情
    </h2>
    """, unsafe_allow_html=True)

    # 检查是否有员工数据
    if stats_df is not None and name_col is not None:
        # 获取所有员工数据（排除"合计"行）
        employees_df = stats_df[stats_df[name_col] != '合计'].copy()

        if not employees_df.empty:
            # 创建员工选择器
            employee_names = sorted(employees_df[name_col].unique().tolist())
            selected_employee = st.selectbox(
                "选择员工查看详情",
                employee_names,
                key="employee_detail_selector"
            )

            # 获取选定员工的数据
            employee_data = employees_df[employees_df[name_col] == selected_employee].iloc[0]

            # 创建左右两列布局
            left_col, right_col = st.columns(2)

            # 左侧显示数据
            with left_col:
                # 基本指标标题 - 不使用卡片背景
                st.markdown("""
                <h4 style="text-align:center; margin-bottom:1rem; color:#1D1D1F; font-size:1.2rem;">基本指标</h4>
                """, unsafe_allow_html=True)

                # 提取数据值
                try:
                    large_orders = int(employee_data['大单数']) if pd.notna(employee_data['大单数']) else 0
                except:
                    large_orders = "数据缺失"

                try:
                    small_orders = int(employee_data['小单数']) if pd.notna(employee_data['小单数']) else 0
                except:
                    small_orders = "数据缺失"

                try:
                    target_small_orders = int(
                        employee_data['目标小单数']) if '目标小单数' in employee_data and pd.notna(
                        employee_data['目标小单数']) else 0
                except:
                    target_small_orders = "数据缺失"

                try:
                    total_amount = float(employee_data['总采购金额']) if pd.notna(employee_data['总采购金额']) else 0
                    total_amount_str = f"¥{total_amount:,.2f}"
                except:
                    total_amount_str = "数据缺失"

                try:
                    target_amount = float(
                        employee_data['本月采购目标额']) if '本月采购目标额' in employee_data and pd.notna(
                        employee_data['本月采购目标额']) else 0
                    target_amount_str = f"¥{target_amount:,.2f}"
                except:
                    target_amount_str = "数据缺失"

                # 创建基本指标的HTML卡片
                basic_metrics_content = f"""
                <div style="display: flex; flex-direction: column; height: 260px; justify-content: center; max-width: 100%;">
                    <div style="display: flex; flex-wrap: wrap; justify-content: space-between; margin-bottom: 2rem; width: 100%;">
                        <div style="flex: 1; min-width: 30%; text-align: center; padding: 0.75rem;">
                            <div style="font-size: 0.9rem; color: var(--color-text-secondary); margin-bottom: 0.5rem;">本月大单数</div>
                            <div style="font-size: 1.6rem; font-weight: bold; color: var(--color-text-primary); white-space: nowrap;">{large_orders}</div>
                        </div>
                        <div style="flex: 1; min-width: 30%; text-align: center; padding: 0.75rem;">
                            <div style="font-size: 0.9rem; color: var(--color-text-secondary); margin-bottom: 0.5rem;">本月小单数</div>
                            <div style="font-size: 1.6rem; font-weight: bold; color: var(--color-text-primary); white-space: nowrap;">{small_orders}</div>
                        </div>
                        <div style="flex: 1; min-width: 30%; text-align: center; padding: 0.75rem;">
                            <div style="font-size: 0.9rem; color: var(--color-text-secondary); margin-bottom: 0.5rem;">目标小单数</div>
                            <div style="font-size: 1.6rem; font-weight: bold; color: var(--color-text-primary); white-space: nowrap;">{target_small_orders}</div>
                        </div>
                    </div>
                    <div style="display: flex; flex-wrap: wrap; justify-content: space-around; margin-top: 1rem; width: 100%;">
                        <div style="flex: 1; min-width: 45%; text-align: center; padding: 0.75rem;">
                            <div style="font-size: 0.9rem; color: var(--color-text-secondary); margin-bottom: 0.5rem;">总采购金额</div>
                            <div style="font-size: 1.6rem; font-weight: bold; color: var(--color-text-primary); white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{total_amount_str}</div>
                        </div>
                        <div style="flex: 1; min-width: 45%; text-align: center; padding: 0.75rem;">
                            <div style="font-size: 0.9rem; color: var(--color-text-secondary); margin-bottom: 0.5rem;">本月采购目标额</div>
                            <div style="font-size: 1.6rem; font-weight: bold; color: var(--color-text-primary); white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{target_amount_str}</div>
                        </div>
                    </div>
                </div>
                """

                # 使用通用卡片组件包装基本指标内容
                basic_metrics_html = create_glass_card(basic_metrics_content, padding="1rem")

                # 渲染基本指标HTML卡片
                components.html(basic_metrics_html, height=320)

                # 周数据标题 - 不使用卡片背景
                st.markdown("""
                <h4 style="text-align:center; margin-bottom:1rem; color:#1D1D1F; font-size:1.2rem;">周数据</h4>
                """, unsafe_allow_html=True)

                # 创建周数据HTML表格
                if week_amount_cols and week_order_cols:
                    # 准备数据
                    weeks = [f"第{i + 1}周" for i in range(len(week_amount_cols))]
                    week_orders = []
                    week_amounts = []

                    # 提取周小单数据
                    for col in week_order_cols:
                        try:
                            value = employee_data[col]
                            if pd.notna(value):
                                if isinstance(value, str) and value.strip() and value.replace('.', '', 1).isdigit():
                                    value = float(value)
                                week_orders.append(int(value))
                            else:
                                week_orders.append(0)
                        except:
                            week_orders.append(0)

                    # 提取周采购金额数据
                    for col in week_amount_cols:
                        try:
                            value = employee_data[col]
                            if pd.notna(value):
                                if isinstance(value, str):
                                    value = float(value.replace(',', '').replace('¥', ''))
                                week_amounts.append(f"¥{value:,.2f}")
                            else:
                                week_amounts.append("¥0.00")
                        except:
                            week_amounts.append("¥0.00")

                    # 生成表格HTML
                    table_rows = ""
                    for i in range(len(weeks)):
                        table_rows += f"""
                        <tr>
                            <td style="padding: 8px; text-align: center; border-bottom: 1px solid rgba(0,0,0,0.05);">{weeks[i]}</td>
                            <td style="padding: 8px; text-align: center; border-bottom: 1px solid rgba(0,0,0,0.05);">{week_orders[i]}</td>
                            <td style="padding: 8px; text-align: center; border-bottom: 1px solid rgba(0,0,0,0.05);">{week_amounts[i]}</td>
                        </tr>
                        """

                    # 表格内容
                    week_data_content = f"""
                    <div style="height: 260px; display: flex; flex-direction: column; justify-content: center; max-width: 100%;">
                        <table style="width: 100%; border-collapse: collapse; font-family: 'SF Pro Text', -apple-system, BlinkMacSystemFont, sans-serif; table-layout: fixed;">
                            <thead>
                                <tr>
                                    <th style="padding: 10px; text-align: center; border-bottom: 2px solid rgba(0,0,0,0.1); width: 20%;">周次</th>
                                    <th style="padding: 10px; text-align: center; border-bottom: 2px solid rgba(0,0,0,0.1); width: 30%;">小单数</th>
                                    <th style="padding: 10px; text-align: center; border-bottom: 2px solid rgba(0,0,0,0.1); width: 50%;">采购金额</th>
                                </tr>
                            </thead>
                            <tbody>
                                {table_rows}
                            </tbody>
                        </table>
                    </div>
                    """

                    # 使用通用卡片组件包装周数据内容
                    week_data_html = create_glass_card(week_data_content, padding="1rem")

                    # 渲染周数据HTML表格
                    components.html(week_data_html, height=320)
                else:
                    # 如果没有周数据，显示提示信息
                    no_data_content = """
                    <div style="height: 260px; display: flex; align-items: center; justify-content: center; width: 100%;">
                        <div style="color: var(--color-text-secondary); font-size: 1rem; text-align: center;">未找到周数据</div>
                    </div>
                    """

                    # 使用通用卡片组件包装无数据提示
                    no_data_html = create_glass_card(no_data_content, padding="1rem")

                    # 渲染无数据提示
                    components.html(no_data_html, height=320)

            # 右侧显示半圆进度图
            with right_col:
                # 小单完成进度标题 - 不使用卡片背景
                st.markdown("""
                <h4 style="text-align:center; margin-bottom:1rem; color:#1D1D1F; font-size:1.2rem;">本月小单完成进度</h4>
                """, unsafe_allow_html=True)

                # 计算小单完成进度
                small_orders_progress = 0
                try:
                    if '目标小单数' in employee_data and pd.notna(employee_data['目标小单数']) and float(
                            employee_data['目标小单数']) > 0:
                        target_small_orders = float(employee_data['目标小单数'])
                        small_orders = float(employee_data['小单数']) if pd.notna(employee_data['小单数']) else 0
                        small_orders_progress = (small_orders / target_small_orders) * 100
                    else:
                        small_orders_progress = 0
                except:
                    small_orders_progress = 0

                # 根据完成进度设置颜色
                if small_orders_progress < 66:
                    progress_color = "#FF453A"  # 红色
                    background_color = "rgba(255, 69, 58, 0.15)"
                elif small_orders_progress < 100:
                    progress_color = "#FFD60A"  # 橙色
                    background_color = "rgba(255, 214, 10, 0.15)"
                else:
                    progress_color = "#30D158"  # 绿色
                    background_color = "rgba(48, 209, 88, 0.15)"

                # 计算SVG中所需的坐标值
                small_progress_capped = min(small_orders_progress, 100)
                small_angle_rad = (small_progress_capped / 100 * 180 - 180) * math.pi / 180
                small_pointer_x = 100 + 90 * math.cos(small_angle_rad)
                small_pointer_y = 100 + 90 * math.sin(small_angle_rad)

                # 计算半圆弧的周长(180°弧长)
                semicircle_length = math.pi * 90
                # 根据进度计算显示的弧长
                small_progress_length = semicircle_length * small_progress_capped / 100

                # 创建小单完成进度的SVG内容
                small_orders_svg = f"""
                <div style="position: relative; width: 100%; height: 200px; margin: 0 auto;">
                    <svg viewBox="0 0 200 100" style="width: 100%; max-width: 300px; overflow: visible;">
                        <!-- 完整的背景半圆，渐变填充区域 -->
                        <defs>
                            <linearGradient id="gradient-bg-small" x1="0%" y1="0%" x2="100%" y2="0%">
                                <stop offset="0%" style="stop-color:rgba(255,69,58,0.15)" />
                                <stop offset="66%" style="stop-color:rgba(255,69,58,0.15)" />
                                <stop offset="66%" style="stop-color:rgba(255,214,10,0.15)" />
                                <stop offset="100%" style="stop-color:rgba(255,214,10,0.15)" />
                            </linearGradient>
                        </defs>

                        <!-- 统一的背景半圆 -->
                        <path d="M10,100 A90,90 0 0,1 190,100" 
                              stroke="url(#gradient-bg-small)" 
                              stroke-width="18" 
                              fill="none" />

                        <!-- 100%标记点 -->
                        <circle cx="190" cy="100" r="4" fill="rgba(48,209,88,0.5)" />

                        <!-- 进度条 - 使用dasharray控制显示长度 -->
                        <path d="M10,100 A90,90 0 0,1 190,100" 
                              stroke="{progress_color}" 
                              stroke-width="18" 
                              fill="none" 
                              stroke-dasharray="{small_progress_length} {semicircle_length}" />

                        <!-- 刻度标记 -->
                        <text x="10" y="118" text-anchor="middle" font-size="8" fill="#666">0%</text>
                        <text x="62" y="50" text-anchor="middle" font-size="8" fill="#666">33%</text>
                        <text x="138" y="50" text-anchor="middle" font-size="8" fill="#666">66%</text>
                        <text x="190" y="118" text-anchor="middle" font-size="8" fill="#666">100%</text>

                        <!-- 指针 -->
                        <line x1="100" 
                              y1="100" 
                              x2="{small_pointer_x}" 
                              y2="{small_pointer_y}" 
                              stroke="black" 
                              stroke-width="2" />

                        <!-- 指针圆心 -->
                        <circle cx="100" cy="100" r="4" fill="#333" />
                    </svg>
                </div>
                <div style="margin-top: 1.5rem; font-size: 1.8rem; font-weight: bold; color: {progress_color};">
                    {small_orders_progress:.1f}%
                </div>
                """

                # 使用通用卡片组件包装SVG内容
                small_orders_html = create_glass_card(small_orders_svg, text_align="center")

                # 使用components.html渲染
                components.html(small_orders_html, height=320)

                # 采购业绩完成进度标题
                st.markdown("""
                <h4 style="text-align:center; margin-bottom:1rem; color:#1D1D1F; font-size:1.2rem;">采购业绩完成进度</h4>
                """, unsafe_allow_html=True)

                # 计算采购业绩完成进度
                purchase_progress = 0
                try:
                    # 先检查是否直接有采购业绩完成进度列
                    if '采购业绩完成进度' in employee_data and pd.notna(employee_data['采购业绩完成进度']):
                        progress_value = employee_data['采购业绩完成进度']
                        if isinstance(progress_value, str) and '%' in progress_value:
                            purchase_progress = float(progress_value.replace('%', ''))
                        else:
                            purchase_progress = float(progress_value) * 100
                    # 如果没有，尝试通过总采购金额和本月采购目标额计算
                    elif '本月采购目标额' in employee_data and pd.notna(employee_data['本月采购目标额']) and float(
                            employee_data['本月采购目标额']) > 0:
                        target_amount = float(employee_data['本月采购目标额'])
                        total_amount = float(employee_data['总采购金额']) if pd.notna(
                            employee_data['总采购金额']) else 0
                        purchase_progress = (total_amount / target_amount) * 100
                    else:
                        purchase_progress = 0
                except:
                    purchase_progress = 0

                # 根据完成进度设置颜色
                if purchase_progress < 66:
                    progress_color = "#FF453A"  # 红色
                    background_color = "rgba(255, 69, 58, 0.15)"
                elif purchase_progress < 100:
                    progress_color = "#FFD60A"  # 橙色
                    background_color = "rgba(255, 214, 10, 0.15)"
                else:
                    progress_color = "#30D158"  # 绿色
                    background_color = "rgba(48, 209, 88, 0.15)"

                # 计算SVG中所需的坐标值
                purchase_progress_capped = min(purchase_progress, 100)
                purchase_angle_rad = (purchase_progress_capped / 100 * 180 - 180) * math.pi / 180
                purchase_pointer_x = 100 + 90 * math.cos(purchase_angle_rad)
                purchase_pointer_y = 100 + 90 * math.sin(purchase_angle_rad)

                # 计算半圆弧的周长(180°弧长)
                # 根据进度计算显示的弧长
                purchase_progress_length = semicircle_length * purchase_progress_capped / 100

                # 创建采购业绩完成进度的SVG内容
                purchase_progress_svg = f"""
                <div style="position: relative; width: 100%; height: 200px; margin: 0 auto;">
                    <svg viewBox="0 0 200 100" style="width: 100%; max-width: 300px; overflow: visible;">
                        <!-- 完整的背景半圆，渐变填充区域 -->
                        <defs>
                            <linearGradient id="gradient-bg-purchase" x1="0%" y1="0%" x2="100%" y2="0%">
                                <stop offset="0%" style="stop-color:rgba(255,69,58,0.15)" />
                                <stop offset="66%" style="stop-color:rgba(255,69,58,0.15)" />
                                <stop offset="66%" style="stop-color:rgba(255,214,10,0.15)" />
                                <stop offset="100%" style="stop-color:rgba(255,214,10,0.15)" />
                            </linearGradient>
                        </defs>

                        <!-- 统一的背景半圆 -->
                        <path d="M10,100 A90,90 0 0,1 190,100" 
                              stroke="url(#gradient-bg-purchase)" 
                              stroke-width="18" 
                              fill="none" />

                        <!-- 100%标记点 -->
                        <circle cx="190" cy="100" r="4" fill="rgba(48,209,88,0.5)" />

                        <!-- 进度条 - 使用dasharray控制显示长度 -->
                        <path d="M10,100 A90,90 0 0,1 190,100" 
                              stroke="{progress_color}" 
                              stroke-width="18" 
                              fill="none" 
                              stroke-dasharray="{purchase_progress_length} {semicircle_length}" />

                        <!-- 刻度标记 -->
                        <text x="10" y="118" text-anchor="middle" font-size="8" fill="#666">0%</text>
                        <text x="62" y="50" text-anchor="middle" font-size="8" fill="#666">33%</text>
                        <text x="138" y="50" text-anchor="middle" font-size="8" fill="#666">66%</text>
                        <text x="190" y="118" text-anchor="middle" font-size="8" fill="#666">100%</text>

                        <!-- 指针 -->
                        <line x1="100" 
                              y1="100" 
                              x2="{purchase_pointer_x}" 
                              y2="{purchase_pointer_y}" 
                              stroke="black" 
                              stroke-width="2" />

                        <!-- 指针圆心 -->
                        <circle cx="100" cy="100" r="4" fill="#333" />
                    </svg>
                </div>
                <div style="margin-top: 1.5rem; font-size: 1.8rem; font-weight: bold; color: {progress_color};">
                    {purchase_progress:.1f}%
                </div>
                """

                # 使用通用卡片组件包装SVG内容
                purchase_progress_html = create_glass_card(purchase_progress_svg, text_align="center")

                # 使用components.html渲染
                components.html(purchase_progress_html, height=320)

        else:
            st.warning("未找到员工数据")
    else:
        st.warning("无法获取员工数据")


# 显示拿货统计页面
def show_delivery_stats_page():
    # 删除标题卡片代码

    # 检查是否已加载数据
    if not st.session_state.get('data_loaded', False) or st.session_state['excel_data'] is None:
        show_warning_card("请先上传采购统计表数据")
        return

    # 获取拿货统计数据
    excel_data = st.session_state['excel_data']
    delivery_df = excel_data.get('本月拿货统计')

    if delivery_df is None:
        show_error_card("未能找到本月拿货统计数据")
        return

    # 第一部分：本月拿货汇总标题
    st.markdown("""
    <h2 style="margin: 20px 0 15px 0; padding: 0; font-size: 28px; font-weight: 700; color: #1D1D1F; font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;">
        一、📊 本月拿货汇总
    </h2>
    """, unsafe_allow_html=True)

    # 获取列名
    name_col = '姓名' if '姓名' in delivery_df.columns else ('报价员' if '报价员' in delivery_df.columns else None)

    if name_col and '合计' in delivery_df[name_col].values:
        # 提取合计行数据
        total_row = delivery_df[delivery_df[name_col] == '合计'].copy()

        # 创建2个指标的列容器
        cols = st.columns(2)

        # 1. 本月拿货量
        if '有效拿货量' in total_row.columns:
            value = total_row['有效拿货量'].values[0]
            # 转换为字符串，然后尝试转换为数字
            value_str = str(value)
            try:
                if value_str and value_str.strip() and value_str.replace('.', '', 1).isdigit():
                    value_num = float(value_str)
                    cols[0].metric("本月拿货量", f"{int(value_num)}")
                else:
                    cols[0].metric("本月拿货量", value_str)
            except:
                cols[0].metric("本月拿货量", value_str)
        else:
            cols[0].metric("本月拿货量", "列不存在")

        # 2. 拿货目标完成度
        if '有效拿货量' in total_row.columns and '目标拿货量' in total_row.columns:
            effective_delivery = total_row['有效拿货量'].values[0]
            target_delivery = total_row['目标拿货量'].values[0]

            try:
                # 转换为数字并计算完成度
                effective_delivery_num = float(str(effective_delivery).replace(',', ''))
                target_delivery_num = float(str(target_delivery).replace(',', ''))

                if target_delivery_num > 0:
                    completion_rate = effective_delivery_num / target_delivery_num
                else:
                    completion_rate = 0

                # 根据完成进度设置不同的颜色
                if completion_rate < 0.66:
                    progress_color = "red"
                elif completion_rate < 1.0:
                    progress_color = "orange"
                else:
                    progress_color = "green"

                # 显示带颜色的进度
                cols[1].markdown(
                    f"""
                    <div style="text-align: center;">
                        <p style="font-size: 14px; color: gray; margin-bottom: 0;">拿货目标完成进度</p>
                        <p style="font-size: 24px; font-weight: bold; color: {progress_color}; margin: 0;">{completion_rate:.1%}</p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
            except Exception as e:
                cols[1].metric("拿货目标完成进度", "计算错误")
        else:
            cols[1].metric("拿货目标完成进度", "列不存在")
    else:
        st.warning("未找到合计行数据，无法显示本月拿货汇总")

    # 第二部分：员工拿货量排名标题
    st.markdown("""
    <h2 style="margin: 30px 0 15px 0; padding: 0; font-size: 28px; font-weight: 700; color: #1D1D1F; font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;">
        二、🏆 员工拿货量排名
    </h2>
    """, unsafe_allow_html=True)

    # 排名数据处理和展示
    if name_col and '有效拿货量' in delivery_df.columns:
        # 去除合计行
        employee_df = delivery_df[delivery_df[name_col] != '合计'].copy()

        # 确保有效拿货量是数值型
        employee_df['有效拿货量'] = pd.to_numeric(employee_df['有效拿货量'], errors='coerce')

        # 按有效拿货量降序排序
        employee_df = employee_df.sort_values(by='有效拿货量', ascending=False).reset_index(drop=True)

        # 计算员工数量
        employee_count = len(employee_df)

        # 使用自定义排名表格而不是matplotlib图表，避免中文和emoji显示问题
        st.markdown("""
        <style>
        .custom-rank-table {
            width: 100%;
            border-collapse: collapse;
            font-family: "微软雅黑", "Microsoft YaHei", sans-serif;
        }
        .custom-rank-row {
            display: flex;
            align-items: center;
            margin-bottom: 8px;
            padding: 10px;
            background-color: rgba(255, 255, 255, 0.8);
            border-radius: 10px;
            box-shadow: 0 2px 5px rgba(0, 0, 0, 0.05);
        }
        .custom-rank-row:hover {
            background-color: rgba(255, 255, 255, 0.95);
            transform: translateY(-2px);
            transition: all 0.2s ease;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.08);
        }
        .rank-number {
            width: 32px;
            height: 32px;
            border-radius: 50%;
            color: white;
            font-weight: bold;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-right: 10px;
            flex-shrink: 0;
        }
        .rank-top-1 {
            background: linear-gradient(135deg, #FFD700, #FFA500);
        }
        .rank-top-2 {
            background: linear-gradient(135deg, #C0C0C0, #A0A0A0);
        }
        .rank-top-3 {
            background: linear-gradient(135deg, #CD7F32, #B87333);
        }
        .rank-normal {
            background: linear-gradient(135deg, #0A84FF, #0055c1);
        }
        .rank-bottom {
            background: linear-gradient(135deg, #FF453A, #c13000);
        }
        .rank-name {
            font-weight: 500;
            width: 80px;
            flex-shrink: 0;
            font-size: 16px;
        }
        .rank-bar-container {
            flex-grow: 1;
            height: 24px;
            background: rgba(0, 0, 0, 0.05);
            border-radius: 12px;
            margin: 0 10px;
            position: relative;
            overflow: hidden;
        }
        .rank-bar {
            height: 100%;
            border-radius: 12px;
            position: absolute;
            top: 0;
            left: 0;
        }
        .bar-top-1 {
            background: linear-gradient(90deg, #FFD700, #FFA500);
        }
        .bar-top-2 {
            background: linear-gradient(90deg, #C0C0C0, #A0A0A0);
        }
        .bar-top-3 {
            background: linear-gradient(90deg, #CD7F32, #B87333);
        }
        .bar-normal {
            background: linear-gradient(90deg, #0A84FF, #0055c1);
        }
        .bar-bottom {
            background: linear-gradient(90deg, #FF453A, #c13000);
        }
        .rank-amount {
            width: 100px;
            text-align: right;
            font-weight: bold;
            flex-shrink: 0;
            font-size: 15px;
        }
        .rank-icon {
            margin-left: 10px;
            font-size: 22px;
            width: 30px;
        }
        </style>
        """, unsafe_allow_html=True)

        # 计算最大拿货量，用于条形图百分比
        max_amount = employee_df['有效拿货量'].max()

        # 遍历员工创建各个排名项
        for i, row in employee_df.iterrows():
            rank = i + 1
            name = row[name_col]
            amount = row['有效拿货量']
            percentage = (amount / max_amount) * 100 if max_amount > 0 else 0

            # 根据排名确定样式和图标
            if rank == 1:
                rank_class = "rank-top-1"
                bar_class = "bar-top-1"
                icon_html = '<span class="rank-icon">🏆</span>'
            elif rank == 2:
                rank_class = "rank-top-2"
                bar_class = "bar-top-2"
                icon_html = '<span class="rank-icon">🥈</span>'
            elif rank == 3:
                rank_class = "rank-top-3"
                bar_class = "bar-top-3"
                icon_html = '<span class="rank-icon">🥉</span>'
            # 修改逻辑：只有当员工数量大于等于7人时，才对最后三名显示警告标记
            elif employee_count >= 7 and employee_count - rank < 3:  # 最后三名（仅当员工数量大于等于7人时）
                rank_class = "rank-bottom"
                bar_class = "bar-bottom"
                icon_html = '<span class="rank-icon">⚠️</span>'
            else:
                rank_class = "rank-normal"
                bar_class = "bar-normal"
                icon_html = '<span class="rank-icon"></span>'  # 空白占位符，保持布局一致

            # 单独渲染每一行
            st.markdown(f"""
            <div class="custom-rank-row">
                <div class="rank-number {rank_class}">{rank}</div>
                <div class="rank-name">{name}</div>
                <div class="rank-bar-container">
                    <div class="rank-bar {bar_class}" style="width: {percentage}%;"></div>
                </div>
                <div class="rank-amount">{int(amount)}</div>
                {icon_html}
            </div>
            """, unsafe_allow_html=True)
    else:
        st.warning("无法显示员工拿货量排名：缺少必要的列或数据")

    # 第三部分标题：周拿货数据分析
    st.markdown("""
    <h2 style="margin: 30px 0 15px 0; padding: 0; font-size: 28px; font-weight: 700; color: #1D1D1F; font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;">
        三、📈 周拿货数据分析
    </h2>
    """, unsafe_allow_html=True)

    # 检查是否有周度数据列
    if name_col and '合计' in delivery_df[name_col].values:
        # 从合计行提取数据
        total_row = delivery_df[delivery_df[name_col] == '合计'].copy().reset_index(drop=True)

        # 初始化周度数据列表
        week_cols_delivery = [col for col in total_row.columns if '周有效拿货量' in col]

        # 提取周度数据
        if week_cols_delivery:
            # 创建周次标签
            weeks = [f"第{i + 1}周" for i in range(len(week_cols_delivery))]

            # 提取周拿货量数据
            week_delivery_amounts = []
            for col in week_cols_delivery:
                try:
                    value = total_row[col].values[0]
                    if pd.notna(value):
                        # 处理可能的字符串格式
                        if isinstance(value, str):
                            value = float(value.replace(',', ''))
                        week_delivery_amounts.append(value)
                    else:
                        week_delivery_amounts.append(0)
                except:
                    week_delivery_amounts.append(0)

            # 创建DataFrame用于表格显示 - 周拿货量
            week_delivery_df = pd.DataFrame({
                '周次': weeks,
                '拿货量': [int(amount) for amount in week_delivery_amounts],
            })

            # 使用Plotly创建交互式图表
            import plotly.graph_objects as go

            # 周拿货量折线图
            fig_delivery = go.Figure()

            # 添加拿货量折线图
            fig_delivery.add_trace(
                go.Scatter(
                    x=weeks,
                    y=week_delivery_amounts,
                    name="周有效拿货量",
                    line=dict(color="#FF9F0A", width=3),
                    mode='lines+markers',
                    marker=dict(size=10, symbol='circle', color="#FF9F0A",
                                line=dict(color='white', width=2))
                )
            )

            # 添加图表样式
            fig_delivery.update_layout(
                title_text="周拿货量趋势分析",
                font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif",
                          size=14),
                plot_bgcolor='rgba(255, 255, 255, 0.8)',  # 设置绘图区背景为白色
                paper_bgcolor='white',  # 设置图表背景为白色
                hovermode="x unified",
                margin=dict(l=30, r=30, t=100, b=30),  # 增加顶部边距
                # 添加边框效果
                shapes=[
                    # 外框
                    dict(
                        type='rect',
                        xref='paper', yref='paper',
                        x0=-0.05, y0=-0.05, x1=1.05, y1=1.05,
                        line=dict(color='rgba(0, 0, 0, 0.08)', width=1),
                        fillcolor='rgba(0, 0, 0, 0)',
                        layer='below'
                    )
                ],
                # 设置图表大小
                width=None,  # 自适应宽度
                height=600,  # 高度
                # 添加图表标题样式
                title={
                    'y': 0.95,
                    'x': 0.5,
                    'xanchor': 'center',
                    'yanchor': 'top',
                    'font': dict(size=18, color="#1D1D1F")
                }
            )

            # 更新轴标签和样式
            fig_delivery.update_xaxes(
                title_text="周次",
                gridcolor='rgba(0, 0, 0, 0.05)',
                showline=False,  # 取消轴边框线
                linewidth=1,
                linecolor='rgba(0, 0, 0, 0.1)',
                tickfont=dict(family="SF Pro Display", size=12)
            )

            # 更新Y轴 (拿货量)
            fig_delivery.update_yaxes(
                title_text="有效拿货量",
                gridcolor='rgba(0, 0, 0, 0.05)',
                showline=False,  # 取消轴边框线
                linewidth=1,
                linecolor='rgba(0, 0, 0, 0.1)',
                tickformat=",.0f"
            )

            # 显示图表
            st.plotly_chart(fig_delivery, use_container_width=True)

            # 显示拿货量表格数据
            st.markdown("<h4 style='text-align: center; margin: 10px 0;'>周度拿货量明细</h4>",
                        unsafe_allow_html=True)
            st.dataframe(week_delivery_df, use_container_width=True, hide_index=True)
        else:
            st.warning("未找到周有效拿货量数据列，无法显示周拿货数据分析")
    else:
        st.warning("未找到合计行数据，无法显示周拿货数据分析")

    # 第四部分标题：员工周拿货数据分析
    st.markdown("""
    <h2 style="margin: 30px 0 15px 0; padding: 0; font-size: 28px; font-weight: 700; color: #1D1D1F; font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;">
        四、👨‍💼 员工周拿货数据分析
    </h2>
    """, unsafe_allow_html=True)

    # 检查是否有周度数据列
    if delivery_df is not None and name_col is not None:
        # 获取所有员工数据（排除"合计"行）
        employees_df = delivery_df[delivery_df[name_col] != '合计'].copy()

        if not employees_df.empty:
            # 找出所有周有效拿货量的列
            week_delivery_cols = [col for col in employees_df.columns if '周有效拿货量' in col]

            if week_delivery_cols:
                # 创建周次标签
                weeks = [f"第{i + 1}周" for i in range(len(week_delivery_cols))]

                # 员工周拿货量折线图
                fig_emp_delivery = go.Figure()

                # 为每个员工添加一条折线
                for i, row in employees_df.iterrows():
                    employee_name = row[name_col]
                    # 提取该员工的周拿货量数据
                    emp_delivery = []
                    for col in week_delivery_cols:
                        try:
                            value = row[col]
                            if pd.notna(value):
                                # 处理可能的字符串格式
                                if isinstance(value, str):
                                    value = float(value.replace(',', ''))
                                emp_delivery.append(value)
                            else:
                                emp_delivery.append(0)
                        except:
                            emp_delivery.append(0)

                    # 添加该员工的折线
                    fig_emp_delivery.add_trace(
                        go.Scatter(
                            x=weeks,
                            y=emp_delivery,
                            name=employee_name,
                            mode='lines+markers',
                            marker=dict(size=8)
                        )
                    )

                # 添加图表样式
                fig_emp_delivery.update_layout(
                    title_text="员工周拿货量趋势对比<br><span style='font-size:12px;color:#86868B'>点击员工图例，隐藏对应数据线</span>",
                    font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif", size=14),
                    plot_bgcolor='rgba(255, 255, 255, 0.8)',
                    paper_bgcolor='white',
                    hovermode="x unified",
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.05,
                        xanchor="center",
                        x=0.485
                    ),
                    margin=dict(l=30, r=30, t=130, b=30),
                    shapes=[
                        dict(
                            type='rect',
                            xref='paper', yref='paper',
                            x0=-0.05, y0=-0.05, x1=1.05, y1=1.05,
                            line=dict(color='rgba(0, 0, 0, 0.08)', width=1),
                            fillcolor='rgba(0, 0, 0, 0)',
                            layer='below'
                        )
                    ],
                    width=None,
                    height=750,
                    title={
                        'y': 0.95,
                        'x': 0.5,
                        'xanchor': 'center',
                        'yanchor': 'top',
                        'font': dict(size=18, color="#1D1D1F")
                    }
                )

                # 更新轴标签和样式
                fig_emp_delivery.update_xaxes(
                    title_text="周次",
                    gridcolor='rgba(0, 0, 0, 0.05)',
                    showline=False,
                    linewidth=1,
                    linecolor='rgba(0, 0, 0, 0.1)',
                    tickfont=dict(family="SF Pro Display", size=12)
                )

                fig_emp_delivery.update_yaxes(
                    title_text="拿货量",
                    gridcolor='rgba(0, 0, 0, 0.05)',
                    showline=False,
                    linewidth=1,
                    linecolor='rgba(0, 0, 0, 0.1)',
                    tickformat=",d"
                )

                # 显示图表
                st.plotly_chart(fig_emp_delivery, use_container_width=True)

                # 创建并显示员工周拿货量详细表格
                emp_delivery_data = []
                for i, row in employees_df.iterrows():
                    emp_row = {name_col: row[name_col]}
                    for j, col in enumerate(week_delivery_cols):
                        try:
                            value = row[col]
                            if pd.notna(value):
                                if isinstance(value, str):
                                    value = float(value.replace(',', ''))
                                emp_row[weeks[j]] = int(value)
                            else:
                                emp_row[weeks[j]] = 0
                        except:
                            emp_row[weeks[j]] = 0
                    emp_delivery_data.append(emp_row)

                emp_delivery_df = pd.DataFrame(emp_delivery_data)
                st.markdown("<h4 style='text-align: center; margin: 10px 0;'>员工周拿货量明细</h4>",
                            unsafe_allow_html=True)
                st.dataframe(emp_delivery_df, use_container_width=True, hide_index=True)
            else:
                st.warning("未找到周有效拿货量数据列，无法显示员工周拿货数据分析")
        else:
            st.warning("未找到员工数据，无法显示员工周拿货数据分析")
    else:
        st.warning("无法获取员工数据")

    # 第五部分标题：员工拿货详情
    st.markdown("""
    <h2 style="margin: 30px 0 15px 0; padding: 0; font-size: 28px; font-weight: 700; color: #1D1D1F; font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;">
        五、👨‍💼 员工拿货详情
    </h2>
    """, unsafe_allow_html=True)

    # 检查是否有员工数据
    if delivery_df is not None and name_col is not None:
        # 获取所有员工数据（排除"合计"行）
        employees_df = delivery_df[delivery_df[name_col] != '合计'].copy()

        if not employees_df.empty:
            # 创建员工选择器
            employee_names = sorted(employees_df[name_col].unique().tolist())
            selected_employee = st.selectbox(
                "选择员工查看拿货详情",
                employee_names,
                key="employee_delivery_selector"
            )

            # 获取选定员工的数据
            employee_data = employees_df[employees_df[name_col] == selected_employee].iloc[0]

            # 创建左右两列布局
            left_col, right_col = st.columns(2)

            # 左侧显示数据
            with left_col:
                # 基本指标标题
                st.markdown("""
                <h4 style="text-align:center; margin: 0.5rem 0 1rem 0; color:#1D1D1F; font-size:1.2rem; height: 1.5rem; line-height: 1.5rem;">拿货指标</h4>
                """, unsafe_allow_html=True)

                # 提取数据值
                try:
                    effective_delivery = float(employee_data['有效拿货量']) if pd.notna(
                        employee_data['有效拿货量']) else 0
                    effective_delivery_str = str(int(effective_delivery))
                except:
                    effective_delivery_str = "数据缺失"

                try:
                    target_delivery = float(employee_data['目标拿货量']) if '目标拿货量' in employee_data and pd.notna(
                        employee_data['目标拿货量']) else 0
                    target_delivery_str = str(int(target_delivery))
                except:
                    target_delivery_str = "数据缺失"

                # 创建基本指标的HTML卡片
                basic_metrics_content = f"""
                <div style="display: flex; flex-direction: column; height: 260px; justify-content: center; max-width: 100%;">
                    <div style="display: flex; flex-wrap: wrap; justify-content: space-around; width: 100%;">
                        <div style="flex: 1; min-width: 45%; text-align: center; padding: 1.5rem;">
                            <div style="font-size: 0.9rem; color: var(--color-text-secondary); margin-bottom: 0.5rem;">本月有效拿货量</div>
                            <div style="font-size: 1.8rem; font-weight: bold; color: var(--color-text-primary); white-space: nowrap;">{effective_delivery_str}</div>
                        </div>
                        <div style="flex: 1; min-width: 45%; text-align: center; padding: 1.5rem;">
                            <div style="font-size: 0.9rem; color: var(--color-text-secondary); margin-bottom: 0.5rem;">本月拿货目标量</div>
                            <div style="font-size: 1.8rem; font-weight: bold; color: var(--color-text-primary); white-space: nowrap;">{target_delivery_str}</div>
                        </div>
                    </div>
                </div>
                """

                # 使用通用卡片组件包装基本指标内容
                basic_metrics_html = create_glass_card(basic_metrics_content, padding="1rem", margin="0")

                # 渲染基本指标HTML卡片
                components.html(basic_metrics_html, height=320)

            # 右侧显示半圆进度图
            with right_col:
                # 拿货完成进度标题
                st.markdown("""
                <h4 style="text-align:center; margin: 0.5rem 0 1rem 0; color:#1D1D1F; font-size:1.2rem; height: 1.5rem; line-height: 1.5rem;">拿货目标完成进度</h4>
                """, unsafe_allow_html=True)

                # 计算拿货完成进度
                delivery_progress = 0
                try:
                    if '目标拿货量' in employee_data and pd.notna(employee_data['目标拿货量']) and float(
                            employee_data['目标拿货量']) > 0:
                        target_delivery = float(employee_data['目标拿货量'])
                        effective_delivery = float(employee_data['有效拿货量']) if pd.notna(
                            employee_data['有效拿货量']) else 0
                        delivery_progress = (effective_delivery / target_delivery) * 100
                    else:
                        delivery_progress = 0
                except:
                    delivery_progress = 0

                # 根据完成进度设置颜色
                if delivery_progress < 66:
                    progress_color = "#FF453A"  # 红色
                    background_color = "rgba(255, 69, 58, 0.15)"
                elif delivery_progress < 100:
                    progress_color = "#FFD60A"  # 橙色
                    background_color = "rgba(255, 214, 10, 0.15)"
                else:
                    progress_color = "#30D158"  # 绿色
                    background_color = "rgba(48, 209, 88, 0.15)"

                # 计算SVG中所需的坐标值
                progress_capped = min(delivery_progress, 100)
                angle_rad = (progress_capped / 100 * 180 - 180) * math.pi / 180
                pointer_x = 100 + 90 * math.cos(angle_rad)
                pointer_y = 100 + 90 * math.sin(angle_rad)

                # 计算半圆弧的周长(180°弧长)
                semicircle_length = math.pi * 90
                # 根据进度计算显示的弧长
                progress_length = semicircle_length * progress_capped / 100

                # 创建拿货完成进度的SVG内容
                delivery_progress_svg = f"""
                <div style="position: relative; width: 100%; height: 200px; margin: 0 auto;">
                    <svg viewBox="0 0 200 100" style="width: 100%; max-width: 300px; overflow: visible;">
                        <!-- 完整的背景半圆，渐变填充区域 -->
                        <defs>
                            <linearGradient id="gradient-bg-delivery" x1="0%" y1="0%" x2="100%" y2="0%">
                                <stop offset="0%" style="stop-color:rgba(255,69,58,0.15)" />
                                <stop offset="66%" style="stop-color:rgba(255,69,58,0.15)" />
                                <stop offset="66%" style="stop-color:rgba(255,214,10,0.15)" />
                                <stop offset="100%" style="stop-color:rgba(255,214,10,0.15)" />
                            </linearGradient>
                        </defs>

                        <!-- 统一的背景半圆 -->
                        <path d="M10,100 A90,90 0 0,1 190,100" 
                              stroke="url(#gradient-bg-delivery)" 
                              stroke-width="18" 
                              fill="none" />

                        <!-- 100%标记点 -->
                        <circle cx="190" cy="100" r="4" fill="rgba(48,209,88,0.5)" />

                        <!-- 进度条 - 使用dasharray控制显示长度 -->
                        <path d="M10,100 A90,90 0 0,1 190,100" 
                              stroke="{progress_color}" 
                              stroke-width="18" 
                              fill="none" 
                              stroke-dasharray="{progress_length} {semicircle_length}" />

                        <!-- 刻度标记 -->
                        <text x="10" y="118" text-anchor="middle" font-size="8" fill="#666">0%</text>
                        <text x="62" y="50" text-anchor="middle" font-size="8" fill="#666">33%</text>
                        <text x="138" y="50" text-anchor="middle" font-size="8" fill="#666">66%</text>
                        <text x="190" y="118" text-anchor="middle" font-size="8" fill="#666">100%</text>

                        <!-- 指针 -->
                        <line x1="100" 
                              y1="100" 
                              x2="{pointer_x}" 
                              y2="{pointer_y}" 
                              stroke="black" 
                              stroke-width="2" />

                        <!-- 指针圆心 -->
                        <circle cx="100" cy="100" r="4" fill="#333" />
                    </svg>
                </div>
                <div style="margin-top: 1.5rem; font-size: 1.8rem; font-weight: bold; color: {progress_color};">
                    {delivery_progress:.1f}%
                </div>
                """

                # 使用通用卡片组件包装SVG内容
                delivery_progress_html = create_glass_card(delivery_progress_svg, text_align="center", margin="0")

                # 使用components.html渲染
                components.html(delivery_progress_html, height=320)
        else:
            st.warning("未找到员工数据，无法显示员工拿货详情")
    else:
        st.warning("无法获取员工数据")


# 显示历月采购拿货对比页面
def show_history_compare_page():
    title_card("历月采购拿货对比")

    # 初始化历月文件会话状态
    if 'history_files' not in st.session_state:
        st.session_state['history_files'] = []
    if 'history_data' not in st.session_state:
        st.session_state['history_data'] = {}
    # 初始化文件上传组件的key，用于重置上传组件状态
    if 'history_uploader_key' not in st.session_state:
        st.session_state['history_uploader_key'] = "history_uploader_1"
    # 初始化文件处理标记，避免重复处理
    if 'processed_history_files' not in st.session_state:
        st.session_state['processed_history_files'] = set()
    # 初始化历月对比子页面状态
    if 'history_compare_subpage' not in st.session_state:
        st.session_state['history_compare_subpage'] = None
    if 'history_compare_params' not in st.session_state:
        st.session_state['history_compare_params'] = {}

    # 创建左右分栏，左侧占1/3，右侧占2/3
    col1, col2 = st.columns([1, 2])

    # 左侧文件上传区
    with col1:
        # 精美的文件上传卡片内容
        upload_content = '''
        <div class="upload-icon">
            <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                <polyline points="17 8 12 3 7 8"></polyline>
                <line x1="12" y1="3" x2="12" y2="15"></line>
            </svg>
        </div>
        <h3 class="upload-title">历月数据上传</h3>
        <div class="upload-description">上传多个采购月度统计表进行对比</div>
        '''

        # 使用create_glass_card函数创建上传卡片
        upload_card_html = create_glass_card(upload_content, text_align="center", padding="2rem")
        components.html(upload_card_html, height=260)

        # 使用动态key创建多文件上传控件，便于重置状态
        uploaded_files = st.file_uploader(
            "选择多个Excel文件",
            type=["xlsx"],
            help="请上传包含'采购详情'、'采购统计'和'本月拿货统计'工作表的Excel文件",
            label_visibility="collapsed",
            accept_multiple_files=True,
            key=st.session_state['history_uploader_key']
        )

        # 处理上传的文件
        if uploaded_files:
            with st.spinner("正在处理文件..."):
                for uploaded_file in uploaded_files:
                    # 检查文件是否已处理，避免重复处理
                    if uploaded_file.name not in st.session_state['processed_history_files']:
                        try:
                            # 保存当前状态到撤销历史
                            save_current_state()

                            # 使用参数target="history"确保数据不会存入主会话状态
                            excel_data, success_msg, file_name = load_excel_data(uploaded_file, target="history")
                            if excel_data is None:
                                show_error_card(f"文件加载失败: {file_name}")
                            else:
                                # 将文件信息和数据添加到历月文件列表
                                st.session_state['history_files'].append({
                                    'name': file_name,
                                    'data': excel_data
                                })
                                st.session_state['history_data'][file_name] = excel_data
                                # 标记文件已处理
                                st.session_state['processed_history_files'].add(file_name)

                                # 使用新的卡片样式显示成功消息
                                success_content = f"""
                                <div style="display: flex; align-items: center; justify-content: center; padding: 0.5rem;">
                                    <span style="font-size: 1.5rem; margin-right: 0.5rem;">✅</span>
                                    <span style="font-size: 1rem; color: var(--color-text-success);">{success_msg}</span>
                                </div>
                                <div style="text-align: center; color: var(--color-text-secondary); margin-top: 0.5rem;">
                                    已加载: {file_name}
                                </div>
                                """
                                success_html = create_glass_card(success_content, text_align="center")
                                components.html(success_html, height=120)
                        except Exception as e:
                            show_error_card(f"处理文件时出错: {str(e)}")

        # 显示已上传的文件列表
        if st.session_state['history_files']:
            st.markdown("### 已上传文件")

            # 为每个文件创建多选框
            selected_files = []
            for idx, file_info in enumerate(st.session_state['history_files']):
                selected = st.checkbox(f"{file_info['name']}", key=f"history_file_{idx}")
                if selected:
                    selected_files.append(idx)

            # 清除按钮区域
            col_btn1, col_btn2 = st.columns(2)

            # 清除选定文件按钮
            with col_btn1:
                if st.button("清除选定文件", use_container_width=True, key="btn_clear_selected"):
                    if selected_files:
                        # 保存当前状态到撤销历史
                        save_current_state()

                        # 从大到小排序索引，以便正确删除
                        for idx in sorted(selected_files, reverse=True):
                            file_name = st.session_state['history_files'][idx]['name']
                            # 从已处理文件集合中移除
                            if file_name in st.session_state['processed_history_files']:
                                st.session_state['processed_history_files'].remove(file_name)
                            # 删除文件数据
                            if file_name in st.session_state['history_data']:
                                del st.session_state['history_data'][file_name]
                            # 删除文件信息
                            st.session_state['history_files'].pop(idx)

                        # 重置上传组件的key，强制重新创建上传组件
                        st.session_state['history_uploader_key'] = f"history_uploader_{hash(str(time.time()))}"
                        st.rerun()
                    else:
                        show_info_card("请先选择需要清除的文件")

            # 清除全部文件按钮
            with col_btn2:
                if st.button("清除全部文件", use_container_width=True, key="btn_clear_all"):
                    # 保存当前状态到撤销历史
                    save_current_state()

                    # 清空历月对比相关的所有状态数据
                    st.session_state['history_files'] = []
                    st.session_state['history_data'] = {}
                    st.session_state['processed_history_files'] = set()
                    # 重置上传组件的key，强制重新创建上传组件
                    st.session_state['history_uploader_key'] = f"history_uploader_{hash(str(time.time()))}"
                    st.rerun()

    # 右侧内容区
    with col2:
        # 功能导航卡片的标题部分
        menu_header = '''
        <div class="menu-header">
            <div class="menu-title-icon">
                <svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M3 3h18v18H3z"></path>
                    <path d="M9 3v18"></path>
                    <path d="M15 3v18"></path>
                    <path d="M3 9h18"></path>
                    <path d="M3 15h18"></path>
                </svg>
            </div>
            <div class="menu-title-text">历月数据对比</div>
        </div>
        <div class="menu-subtitle">选择对比分析类型</div>
        <div class="menu-divider"></div>
        '''

        # 使用create_glass_card函数创建导航卡片头部
        menu_card_html = create_glass_card(menu_header, padding="1.5rem 1.5rem 0 1.5rem")
        components.html(menu_card_html, height=180)

        # 创建两个功能按钮
        compare_items = [
            {
                "icon": '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="1" y="6" width="18" height="12" rx="2"></rect><path d="m4 10 4 2 4-2"></path></svg>',
                "title": "历月采购额对比",
                "description": "对比不同月份的采购金额",
                "key": "btn_history_purchase_compare",
                "subpage": "history_purchase_compare",
                "color": "info"
            },
            {
                "icon": '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 8h2a2 2 0 0 1 2 2v2a2 2 0 0 1-2 2h-2"></path><path d="M3 8h2a2 2 0 0 1 2 2v2a2 2 0 0 1-2 2H3"></path><path d="M10 8h4"></path><path d="M10 12h4"></path><path d="M10 16h4"></path></svg>',
                "title": "历月拿货量对比",
                "description": "对比不同月份的拿货数量",
                "key": "btn_history_delivery_compare",
                "subpage": "history_delivery_compare",
                "color": "warning"
            }
        ]

        # 显示功能按钮，每行一个按钮
        for i, item in enumerate(compare_items):
            # 创建一个简单的按钮，但应用自定义CSS使其看起来像卡片
            if st.button(
                    f"**{item['title']}**\n\n{item['description']}",
                    key=item["key"],
                    use_container_width=True
            ):
                # 保存当前状态到撤销历史
                save_current_state()

                # 设置子页面状态
                st.session_state['history_compare_subpage'] = item["subpage"]
                st.rerun()

            # 添加自定义CSS，为按钮添加图标和样式
            st.markdown(f"""
            <style>
                /* 为按钮 {item['key']} 添加自定义样式 */
                [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"] {{
                    background: var(--glass-bg);
                    border-radius: 16px;
                    padding: 1.5rem 1rem;
                    height: auto;
                    min-height: 100px;
                    display: flex;
                    flex-direction: column;
                    align-items: flex-start;
                    text-align: left;
                    font-family: 'SF Pro Text', sans-serif;
                    transition: all 0.3s cubic-bezier(0.175, 0.885, 0.32, 1.1);
                    border: 1px solid rgba(255, 255, 255, 0.5);
                    position: relative;
                    overflow: hidden;
                    margin-bottom: 1rem;
                }}

                [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"]:hover {{
                    transform: translateY(-5px) scale(1.02);
                    box-shadow: 0 15px 30px rgba(0, 0, 0, 0.08);
                }}

                [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"]:before {{
                    content: '';
                    position: absolute;
                    top: 0.75rem;
                    right: 0.75rem;
                    width: 40px;
                    height: 40px;
                    border-radius: 50%;
                    background-color: var(--color-{item['color']}-soft);
                    opacity: 0.8;
                    z-index: 0;
                }}

                [data-testid="stButton"] > button[kind="secondary"][aria-label="{item['key']}"]:after {{
                    content: '{item['icon']}';
                    position: absolute;
                    top: 0.75rem;
                    right: 0.75rem;
                    width: 40px;
                    height: 40px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    color: var(--color-{item['color']});
                    z-index: 1;
                }}
            </style>
            """, unsafe_allow_html=True)


# 显示历月采购额对比页面
def show_history_purchase_compare():
    title_card("历月采购额对比分析")

    # 检查是否有历月数据
    if not st.session_state.get('history_files', []) or not st.session_state.get('history_data', {}):
        show_warning_card("请先在历月对比页面上传历月数据文件")
        return

    # 第一部分：总采购金额对比
    st.markdown("""
    <h2 style="margin: 20px 0 15px 0; padding: 0; font-size: 28px; font-weight: 700; color: #1D1D1F; font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;">
        一、📊 总采购金额对比
    </h2>
    """, unsafe_allow_html=True)

    # 获取所有上传的文件数据
    history_files = st.session_state['history_files']
    history_data = st.session_state['history_data']

    # 提取每个文件的总采购金额
    purchase_data = []
    for file_info in history_files:
        file_name = file_info['name']
        file_data = history_data.get(file_name)

        if file_data and '采购统计' in file_data:
            stats_df = file_data['采购统计']

            # 获取姓名列
            name_col = '姓名' if '姓名' in stats_df.columns else ('报价员' if '报价员' in stats_df.columns else None)

            if name_col and '合计' in stats_df[name_col].values and '总采购金额' in stats_df.columns:
                # 提取合计行数据
                total_row = stats_df[stats_df[name_col] == '合计'].copy()

                if not total_row.empty:
                    try:
                        # 提取总采购金额
                        amount_value = total_row['总采购金额'].values[0]

                        # 处理可能的字符串格式
                        if isinstance(amount_value, str):
                            amount_value = float(amount_value.replace(',', '').replace('¥', ''))

                        # 尝试从文件名提取年月信息
                        import re
                        match = re.search(r'采购(\d{4})年_(\d{1,2})月统计表', file_name)
                        if match:
                            year = match.group(1)
                            month = match.group(2)
                            month_label = f"{year}年{month}月"
                        else:
                            # 如果无法从文件名提取，则使用文件名作为标签
                            month_label = file_name.replace("采购", "").replace("统计表.xlsx", "")

                        purchase_data.append({
                            'month_label': month_label,
                            'file_name': file_name,
                            'amount': float(amount_value)
                        })
                    except Exception as e:
                        st.warning(f"处理文件 '{file_name}' 时出错: {str(e)}")

    # 如果没有提取到数据
    if not purchase_data:
        show_warning_card("未能从上传的文件中提取到有效的总采购金额数据")
        return

    # 按月份排序（假设月份标签格式为"YYYY年MM月"）
    try:
        purchase_data.sort(key=lambda x: (
            int(x['month_label'].split('年')[0]) if '年' in x['month_label'] else 0,
            int(x['month_label'].split('年')[1].split('月')[0]) if '年' in x['month_label'] and '月' in x[
                'month_label'] else 0
        ))
    except:
        # 如果排序失败，保持原顺序
        pass

    # 提取排序后的月份标签和金额数据
    months = [item['month_label'] for item in purchase_data]
    amounts = [item['amount'] for item in purchase_data]

    # 使用Plotly创建交互式折线图
    import plotly.graph_objects as go

    # 创建总采购金额折线图
    fig_amount = go.Figure()

    # 添加金额折线图
    fig_amount.add_trace(
        go.Scatter(
            x=months,
            y=amounts,
            name="总采购金额",
            line=dict(color="#0A84FF", width=3),
            mode='lines+markers',
            marker=dict(size=10, symbol='circle', color="#0A84FF",
                        line=dict(color='white', width=2))
        )
    )

    # 添加图表样式
    fig_amount.update_layout(
        title_text="历月总采购金额趋势分析",
        font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif", size=14),
        plot_bgcolor='rgba(255, 255, 255, 0.8)',
        paper_bgcolor='white',
        hovermode="x unified",
        margin=dict(l=30, r=30, t=100, b=30),
        shapes=[
            # 外框
            dict(
                type='rect',
                xref='paper', yref='paper',
                x0=-0.05, y0=-0.05, x1=1.05, y1=1.05,
                line=dict(color='rgba(0, 0, 0, 0.08)', width=1),
                fillcolor='rgba(0, 0, 0, 0)',
                layer='below'
            )
        ],
        width=None,
        height=650,
        title={
            'y': 0.95,
            'x': 0.5,
            'xanchor': 'center',
            'yanchor': 'top',
            'font': dict(size=18, color="#1D1D1F")
        },
    )

    # 更新轴标签和样式
    fig_amount.update_xaxes(
        title_text="月份",
        gridcolor='rgba(0, 0, 0, 0.05)',
        showline=False,
        linewidth=1,
        linecolor='rgba(0, 0, 0, 0.1)',
        tickfont=dict(family="SF Pro Display", size=12)
    )

    fig_amount.update_yaxes(
        title_text="采购金额 (元)",
        gridcolor='rgba(0, 0, 0, 0.05)',
        showline=False,
        linewidth=1,
        linecolor='rgba(0, 0, 0, 0.1)',
        tickformat=",.0f"
    )

    # 显示折线图
    st.plotly_chart(fig_amount, use_container_width=True)

    # 创建表格数据
    amount_df = pd.DataFrame({
        '月份': months,
        '总采购金额': [f"¥{amount:,.2f}" for amount in amounts]
    })

    # 显示金额表格
    st.markdown("<h4 style='text-align: center; margin: 10px 0;'>历月总采购金额明细</h4>",
                unsafe_allow_html=True)
    st.dataframe(amount_df, use_container_width=True, hide_index=True)

    # 计算环比增长率
    if len(amounts) > 1:
        growth_rates = []
        for i in range(1, len(amounts)):
            if amounts[i - 1] != 0:
                growth_rate = (amounts[i] - amounts[i - 1]) / amounts[i - 1] * 100
                # 添加正号(+)显示正数增长率
                if growth_rate > 0:
                    growth_rates.append(f"+{growth_rate:.2f}%")
                else:
                    growth_rates.append(f"{growth_rate:.2f}%")
            else:
                growth_rates.append("无可用上月数据")

        # 为第一个月添加"无可用上月数据"
        growth_rates = ["无可用上月数据"] + growth_rates

        # 添加环比增长率到表格
        amount_df['环比增长率'] = growth_rates

        # 显示带环比增长率的表格
        st.markdown("<h4 style='text-align: center; margin: 20px 0 10px 0;'>历月总采购金额环比分析</h4>",
                    unsafe_allow_html=True)

        # 为环比增长率列应用样式，灰色显示"无可用上月数据"
        def highlight_unavailable(val):
            if val == "无可用上月数据":
                return 'color: gray'
            elif "+" in val:
                return 'color: green'
            elif "-" in val:
                return 'color: red'
            return ''

        # 应用样式并显示
        styled_df = amount_df.style.map(highlight_unavailable, subset=['环比增长率'])
        st.dataframe(styled_df, use_container_width=True, hide_index=True)

    # 第二部分：历月小单数统计
    st.markdown("""
    <h2 style="margin: 40px 0 15px 0; padding: 0; font-size: 28px; font-weight: 700; color: #1D1D1F; font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;">
        二、📈 历月小单数统计
    </h2>
    """, unsafe_allow_html=True)

    # 提取每个文件的小单数
    orders_data = []
    for file_info in history_files:
        file_name = file_info['name']
        file_data = history_data.get(file_name)

        if file_data and '采购统计' in file_data:
            stats_df = file_data['采购统计']

            # 获取姓名列
            name_col = '姓名' if '姓名' in stats_df.columns else ('报价员' if '报价员' in stats_df.columns else None)

            if name_col and '合计' in stats_df[name_col].values and '小单数' in stats_df.columns:
                # 提取合计行数据
                total_row = stats_df[stats_df[name_col] == '合计'].copy()

                if not total_row.empty:
                    try:
                        # 提取小单数
                        orders_value = total_row['小单数'].values[0]

                        # 处理可能的字符串格式
                        if isinstance(orders_value, str) and orders_value.strip() and orders_value.replace('.', '',
                                                                                                           1).isdigit():
                            orders_value = float(orders_value)

                        # 尝试从文件名提取年月信息
                        match = re.search(r'采购(\d{4})年_(\d{1,2})月统计表', file_name)
                        if match:
                            year = match.group(1)
                            month = match.group(2)
                            month_label = f"{year}年{month}月"
                        else:
                            # 如果无法从文件名提取，则使用文件名作为标签
                            month_label = file_name.replace("采购", "").replace("统计表.xlsx", "")

                        orders_data.append({
                            'month_label': month_label,
                            'file_name': file_name,
                            'orders': float(orders_value)
                        })
                    except Exception as e:
                        st.warning(f"处理文件 '{file_name}' 小单数数据时出错: {str(e)}")

    # 如果没有提取到数据
    if not orders_data:
        show_warning_card("未能从上传的文件中提取到有效的小单数数据")
        return

    # 按月份排序（与总采购金额部分相同的逻辑）
    try:
        orders_data.sort(key=lambda x: (
            int(x['month_label'].split('年')[0]) if '年' in x['month_label'] else 0,
            int(x['month_label'].split('年')[1].split('月')[0]) if '年' in x['month_label'] and '月' in x[
                'month_label'] else 0
        ))
    except:
        # 如果排序失败，保持原顺序
        pass

    # 提取排序后的月份标签和小单数数据
    orders_months = [item['month_label'] for item in orders_data]
    orders_counts = [item['orders'] for item in orders_data]

    # 创建小单数折线图
    fig_orders = go.Figure()

    # 添加小单数折线图
    fig_orders.add_trace(
        go.Scatter(
            x=orders_months,
            y=orders_counts,
            name="小单数",
            line=dict(color="#30D158", width=3),
            mode='lines+markers',
            marker=dict(size=10, symbol='diamond', color="#30D158",
                        line=dict(color='white', width=2))
        )
    )

    # 添加图表样式
    fig_orders.update_layout(
        title_text="历月小单数趋势分析",
        font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif", size=14),
        plot_bgcolor='rgba(255, 255, 255, 0.8)',
        paper_bgcolor='white',
        hovermode="x unified",
        margin=dict(l=30, r=30, t=100, b=30),
        shapes=[
            # 外框
            dict(
                type='rect',
                xref='paper', yref='paper',
                x0=-0.05, y0=-0.05, x1=1.05, y1=1.05,
                line=dict(color='rgba(0, 0, 0, 0.08)', width=1),
                fillcolor='rgba(0, 0, 0, 0)',
                layer='below'
            )
        ],
        width=None,
        height=650,
        title={
            'y': 0.95,
            'x': 0.5,
            'xanchor': 'center',
            'yanchor': 'top',
            'font': dict(size=18, color="#1D1D1F")
        },
    )

    # 更新轴标签和样式
    fig_orders.update_xaxes(
        title_text="月份",
        gridcolor='rgba(0, 0, 0, 0.05)',
        showline=False,
        linewidth=1,
        linecolor='rgba(0, 0, 0, 0.1)',
        tickfont=dict(family="SF Pro Display", size=12)
    )

    fig_orders.update_yaxes(
        title_text="小单数",
        gridcolor='rgba(0, 0, 0, 0.05)',
        showline=False,
        linewidth=1,
        linecolor='rgba(0, 0, 0, 0.1)',
        tickformat=",.0f"
    )

    # 显示折线图
    st.plotly_chart(fig_orders, use_container_width=True)

    # 创建表格数据
    orders_df = pd.DataFrame({
        '月份': orders_months,
        '小单数': [int(count) for count in orders_counts]
    })

    # 显示表格
    st.markdown("<h4 style='text-align: center; margin: 10px 0;'>历月小单数明细</h4>",
                unsafe_allow_html=True)
    st.dataframe(orders_df, use_container_width=True, hide_index=True)

    # 计算环比增长率
    if len(orders_counts) > 1:
        growth_rates = []
        for i in range(1, len(orders_counts)):
            if orders_counts[i - 1] != 0:
                growth_rate = (orders_counts[i] - orders_counts[i - 1]) / orders_counts[i - 1] * 100
                # 添加正号(+)显示正数增长率
                if growth_rate > 0:
                    growth_rates.append(f"+{growth_rate:.2f}%")
                else:
                    growth_rates.append(f"{growth_rate:.2f}%")
            else:
                growth_rates.append("无可用上月数据")

        # 为第一个月添加"无可用上月数据"
        growth_rates = ["无可用上月数据"] + growth_rates

        # 添加环比增长率到表格
        orders_df['环比增长率'] = growth_rates

        # 显示带环比增长率的表格
        st.markdown("<h4 style='text-align: center; margin: 20px 0 10px 0;'>历月小单数环比分析</h4>",
                    unsafe_allow_html=True)

        # 为环比增长率列应用样式，灰色显示"无可用上月数据"
        def highlight_unavailable(val):
            if val == "无可用上月数据":
                return 'color: gray'
            elif "+" in val:
                return 'color: green'
            elif "-" in val:
                return 'color: red'
            return ''

        # 应用样式并显示
        styled_df = orders_df.style.map(highlight_unavailable, subset=['环比增长率'])
        st.dataframe(styled_df, use_container_width=True, hide_index=True)

    # 第三部分：员工月度采购分析
    st.markdown("""
    <h2 style="margin: 40px 0 15px 0; padding: 0; font-size: 28px; font-weight: 700; color: #1D1D1F; font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;">
        三、👥 员工月度采购分析
    </h2>
    """, unsafe_allow_html=True)

    # 提取所有文件中的员工数据
    all_employees = set()
    employee_data = {}

    # 遍历每个文件
    for file_info in history_files:
        file_name = file_info['name']
        file_data = history_data.get(file_name)

        # 提取月份标签
        match = re.search(r'采购(\d{4})年_(\d{1,2})月统计表', file_name)
        if match:
            year = match.group(1)
            month = match.group(2)
            month_label = f"{year}年{month}月"
        else:
            month_label = file_name.replace("采购", "").replace("统计表.xlsx", "")

        if file_data and '采购统计' in file_data:
            stats_df = file_data['采购统计']

            # 获取姓名列
            name_col = '姓名' if '姓名' in stats_df.columns else ('报价员' if '报价员' in stats_df.columns else None)

            if name_col:
                # 过滤掉合计行
                employee_rows = stats_df[stats_df[name_col] != '合计'].copy()

                for _, row in employee_rows.iterrows():
                    emp_name = row[name_col]
                    if pd.notna(emp_name) and emp_name:
                        # 添加到员工集合
                        all_employees.add(emp_name)

                        # 初始化该员工的数据字典
                        if emp_name not in employee_data:
                            employee_data[emp_name] = {}

                        # 初始化该月份的数据
                        if month_label not in employee_data[emp_name]:
                            employee_data[emp_name][month_label] = {
                                'small_orders': 0,
                                'target_small_orders': 0,  # 添加目标小单数字段
                                'purchase_amount': 0,
                                'small_orders_progress': 0,
                                'purchase_progress': 0
                            }

                        # 提取小单数
                        if '小单数' in row:
                            try:
                                orders_value = row['小单数']
                                if pd.notna(orders_value):
                                    if isinstance(orders_value, str) and orders_value.strip() and orders_value.replace(
                                            '.', '', 1).isdigit():
                                        orders_value = float(orders_value)
                                    employee_data[emp_name][month_label]['small_orders'] = float(orders_value)
                            except:
                                pass

                        # 提取目标小单数
                        if '目标小单数' in row:
                            try:
                                target_orders_value = row['目标小单数']
                                if pd.notna(target_orders_value):
                                    if isinstance(target_orders_value,
                                                  str) and target_orders_value.strip() and target_orders_value.replace(
                                            '.', '', 1).isdigit():
                                        target_orders_value = float(target_orders_value)
                                    employee_data[emp_name][month_label]['target_small_orders'] = float(
                                        target_orders_value)
                            except:
                                pass

                        # 提取总采购金额
                        if '总采购金额' in row:
                            try:
                                amount_value = row['总采购金额']
                                if pd.notna(amount_value):
                                    if isinstance(amount_value, str):
                                        amount_value = float(amount_value.replace(',', '').replace('¥', ''))
                                    employee_data[emp_name][month_label]['purchase_amount'] = float(amount_value)
                            except:
                                pass

                        # 提取采购业绩完成进度
                        if '采购业绩完成进度' in row:
                            try:
                                progress_value = row['采购业绩完成进度']
                                if pd.notna(progress_value):
                                    # 处理百分比字符串情况
                                    if isinstance(progress_value, str) and '%' in progress_value:
                                        progress_value = float(progress_value.replace('%', '')) / 100
                                    employee_data[emp_name][month_label]['purchase_progress'] = float(progress_value)
                            except:
                                pass

    # 计算小单目标完成进度
    for emp_name in employee_data:
        for month_label in employee_data[emp_name]:
            small_orders = employee_data[emp_name][month_label]['small_orders']
            target_small_orders = employee_data[emp_name][month_label]['target_small_orders']

            # 计算小单完成进度 = 小单数/目标小单数
            if target_small_orders > 0:
                employee_data[emp_name][month_label]['small_orders_progress'] = small_orders / target_small_orders
            else:
                employee_data[emp_name][month_label]['small_orders_progress'] = 0

    # 如果未找到员工数据
    if not employee_data:
        show_warning_card("未能从上传的文件中提取到员工数据")
        return

    # 将员工列表转换为有序列表，以便在多选框中显示
    employee_list = sorted(list(all_employees))

    # 员工选择多选框
    st.markdown("<div style='margin: 20px 0;'></div>", unsafe_allow_html=True)
    selected_employees = st.multiselect(
        "选择要分析的员工",
        employee_list,
        default=employee_list[:min(3, len(employee_list))],  # 默认选择前3个员工
        key="employee_selector"
    )

    # 如果没有选择员工，显示提示信息
    if not selected_employees:
        show_info_card("请选择至少一名员工进行分析")
        return

    # 获取所有月份标签并排序
    all_months = set()
    for emp in employee_data:
        all_months.update(employee_data[emp].keys())

    # 将月份列表转换为有序列表
    try:
        month_list = sorted(list(all_months), key=lambda x: (
            int(x.split('年')[0]) if '年' in x else 0,
            int(x.split('年')[1].split('月')[0]) if '年' in x and '月' in x else 0
        ))
    except:
        month_list = list(all_months)

    # 创建左右分栏
    left_col, right_col = st.columns(2)

    # 左侧上方：员工小单数折线图
    with left_col:
        # 准备小单数据
        emp_order_data = []
        for emp in selected_employees:
            if emp in employee_data:
                emp_months = []
                emp_orders = []

                for month in month_list:
                    if month in employee_data[emp]:
                        emp_months.append(month)
                        emp_orders.append(employee_data[emp][month]['small_orders'])

                emp_order_data.append({
                    'name': emp,
                    'months': emp_months,
                    'orders': emp_orders
                })

        # 创建小单数折线图
        fig_emp_orders = go.Figure()

        # 为每个员工添加折线
        colors = ['#0A84FF', '#30D158', '#FF9F0A', '#FF375F', '#5E5CE6', '#BF5AF2', '#AC8E68', '#64D2FF']
        for i, emp_data in enumerate(emp_order_data):
            color_idx = i % len(colors)
            fig_emp_orders.add_trace(
                go.Scatter(
                    x=emp_data['months'],
                    y=emp_data['orders'],
                    name=emp_data['name'],
                    line=dict(color=colors[color_idx], width=2),
                    mode='lines+markers',
                    marker=dict(size=8, symbol='circle', color=colors[color_idx],
                                line=dict(color='white', width=1))
                )
            )

        # 添加图表样式
        fig_emp_orders.update_layout(
            title_text="月度小单趋势",
            font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif", size=14),
            plot_bgcolor='rgba(255, 255, 255, 0.8)',
            paper_bgcolor='white',
            hovermode="x unified",
            margin=dict(l=20, r=20, t=80, b=30),
            shapes=[
                # 外框
                dict(
                    type='rect',
                    xref='paper', yref='paper',
                    x0=-0.05, y0=-0.05, x1=1.05, y1=1.05,
                    line=dict(color='rgba(0, 0, 0, 0.08)', width=1),
                    fillcolor='rgba(0, 0, 0, 0)',
                    layer='below'
                )
            ],
            height=350,
            title={
                'y': 0.95,
                'x': 0.5,
                'xanchor': 'center',
                'yanchor': 'top',
                'font': dict(size=16, color="#1D1D1F")
            },
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
                font=dict(size=11)
            )
        )

        fig_emp_orders.update_xaxes(
            title_text="月份",
            gridcolor='rgba(0, 0, 0, 0.05)',
            showline=False,
            linewidth=1,
            linecolor='rgba(0, 0, 0, 0.1)',
            tickfont=dict(size=11)
        )

        fig_emp_orders.update_yaxes(
            title_text="小单数",
            gridcolor='rgba(0, 0, 0, 0.05)',
            showline=False,
            linewidth=1,
            linecolor='rgba(0, 0, 0, 0.1)',
            tickformat=",.0f",
            tickfont=dict(size=11)
        )

        # 显示小单数折线图
        st.plotly_chart(fig_emp_orders, use_container_width=True)

        # 准备总采购金额数据
        emp_amount_data = []
        for emp in selected_employees:
            if emp in employee_data:
                emp_months = []
                emp_amounts = []

                for month in month_list:
                    if month in employee_data[emp]:
                        emp_months.append(month)
                        emp_amounts.append(employee_data[emp][month]['purchase_amount'])

                emp_amount_data.append({
                    'name': emp,
                    'months': emp_months,
                    'amounts': emp_amounts
                })

        # 创建总采购金额折线图
        fig_emp_amounts = go.Figure()

        # 为每个员工添加折线
        for i, emp_data in enumerate(emp_amount_data):
            color_idx = i % len(colors)
            fig_emp_amounts.add_trace(
                go.Scatter(
                    x=emp_data['months'],
                    y=emp_data['amounts'],
                    name=emp_data['name'],
                    line=dict(color=colors[color_idx], width=2),
                    mode='lines+markers',
                    marker=dict(size=8, symbol='diamond', color=colors[color_idx],
                                line=dict(color='white', width=1))
                )
            )

        # 添加图表样式
        fig_emp_amounts.update_layout(
            title_text="月度采购金额趋势",
            font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif", size=14),
            plot_bgcolor='rgba(255, 255, 255, 0.8)',
            paper_bgcolor='white',
            hovermode="x unified",
            margin=dict(l=20, r=20, t=80, b=30),
            shapes=[
                # 外框
                dict(
                    type='rect',
                    xref='paper', yref='paper',
                    x0=-0.05, y0=-0.05, x1=1.05, y1=1.05,
                    line=dict(color='rgba(0, 0, 0, 0.08)', width=1),
                    fillcolor='rgba(0, 0, 0, 0)',
                    layer='below'
                )
            ],
            height=350,
            title={
                'y': 0.95,
                'x': 0.5,
                'xanchor': 'center',
                'yanchor': 'top',
                'font': dict(size=16, color="#1D1D1F")
            },
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
                font=dict(size=11)
            )
        )

        fig_emp_amounts.update_xaxes(
            title_text="月份",
            gridcolor='rgba(0, 0, 0, 0.05)',
            showline=False,
            linewidth=1,
            linecolor='rgba(0, 0, 0, 0.1)',
            tickfont=dict(size=11)
        )

        fig_emp_amounts.update_yaxes(
            title_text="采购金额 (元)",
            gridcolor='rgba(0, 0, 0, 0.05)',
            showline=False,
            linewidth=1,
            linecolor='rgba(0, 0, 0, 0.1)',
            tickformat=",.0f",
            tickfont=dict(size=11)
        )

        # 显示总采购金额折线图
        st.plotly_chart(fig_emp_amounts, use_container_width=True)

    # 右侧：目标完成度分析
    with right_col:
        # 准备小单目标完成进度数据
        emp_order_progress_data = []
        for emp in selected_employees:
            if emp in employee_data:
                emp_months = []
                emp_progress = []

                for month in month_list:
                    if month in employee_data[emp]:
                        emp_months.append(month)
                        emp_progress.append(employee_data[emp][month]['small_orders_progress'])

                emp_order_progress_data.append({
                    'name': emp,
                    'months': emp_months,
                    'progress': emp_progress
                })

        # 创建小单目标完成进度折线图
        fig_order_progress = go.Figure()

        # 添加参考线（100%完成度）
        fig_order_progress.add_shape(
            type="line",
            x0=0,
            y0=1,
            x1=1,
            y1=1,
            line=dict(
                color="rgba(0, 0, 0, 0.3)",
                width=1,
                dash="dash",
            ),
            xref="paper",
            yref="y"
        )

        # 为每个员工添加折线
        for i, emp_data in enumerate(emp_order_progress_data):
            color_idx = i % len(colors)
            fig_order_progress.add_trace(
                go.Scatter(
                    x=emp_data['months'],
                    y=emp_data['progress'],
                    name=emp_data['name'],
                    line=dict(color=colors[color_idx], width=2),
                    mode='lines+markers',
                    marker=dict(size=8, symbol='circle', color=colors[color_idx],
                                line=dict(color='white', width=1))
                )
            )

        # 添加图表样式
        fig_order_progress.update_layout(
            title_text="小单目标完成率",
            font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif", size=14),
            plot_bgcolor='rgba(255, 255, 255, 0.8)',
            paper_bgcolor='white',
            hovermode="x unified",
            margin=dict(l=20, r=20, t=80, b=30),
            shapes=[
                # 外框
                dict(
                    type='rect',
                    xref='paper', yref='paper',
                    x0=-0.05, y0=-0.05, x1=1.05, y1=1.05,
                    line=dict(color='rgba(0, 0, 0, 0.08)', width=1),
                    fillcolor='rgba(0, 0, 0, 0)',
                    layer='below'
                )
            ],
            height=350,
            title={
                'y': 0.95,
                'x': 0.5,
                'xanchor': 'center',
                'yanchor': 'top',
                'font': dict(size=16, color="#1D1D1F")
            },
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
                font=dict(size=11)
            )
        )

        fig_order_progress.update_xaxes(
            title_text="月份",
            gridcolor='rgba(0, 0, 0, 0.05)',
            showline=False,
            linewidth=1,
            linecolor='rgba(0, 0, 0, 0.1)',
            tickfont=dict(size=11)
        )

        fig_order_progress.update_yaxes(
            title_text="小单目标完成进度",
            gridcolor='rgba(0, 0, 0, 0.05)',
            showline=False,
            linewidth=1,
            linecolor='rgba(0, 0, 0, 0.1)',
            tickformat=".0%",
            tickfont=dict(size=11)
        )

        # 显示小单目标完成进度折线图
        st.plotly_chart(fig_order_progress, use_container_width=True)

        # 准备采购业绩完成进度数据
        emp_purchase_progress_data = []
        for emp in selected_employees:
            if emp in employee_data:
                emp_months = []
                emp_progress = []

                for month in month_list:
                    if month in employee_data[emp]:
                        emp_months.append(month)
                        emp_progress.append(employee_data[emp][month]['purchase_progress'])

                emp_purchase_progress_data.append({
                    'name': emp,
                    'months': emp_months,
                    'progress': emp_progress
                })

        # 创建采购业绩完成进度折线图
        fig_purchase_progress = go.Figure()

        # 添加参考线（100%完成度）
        fig_purchase_progress.add_shape(
            type="line",
            x0=0,
            y0=1,
            x1=1,
            y1=1,
            line=dict(
                color="rgba(0, 0, 0, 0.3)",
                width=1,
                dash="dash",
            ),
            xref="paper",
            yref="y"
        )

        # 为每个员工添加折线
        for i, emp_data in enumerate(emp_purchase_progress_data):
            color_idx = i % len(colors)
            fig_purchase_progress.add_trace(
                go.Scatter(
                    x=emp_data['months'],
                    y=emp_data['progress'],
                    name=emp_data['name'],
                    line=dict(color=colors[color_idx], width=2),
                    mode='lines+markers',
                    marker=dict(size=8, symbol='diamond', color=colors[color_idx],
                                line=dict(color='white', width=1))
                )
            )

        # 添加图表样式
        fig_purchase_progress.update_layout(
            title_text="采购目标完成率",
            font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif", size=14),
            plot_bgcolor='rgba(255, 255, 255, 0.8)',
            paper_bgcolor='white',
            hovermode="x unified",
            margin=dict(l=20, r=20, t=80, b=30),
            shapes=[
                # 外框
                dict(
                    type='rect',
                    xref='paper', yref='paper',
                    x0=-0.05, y0=-0.05, x1=1.05, y1=1.05,
                    line=dict(color='rgba(0, 0, 0, 0.08)', width=1),
                    fillcolor='rgba(0, 0, 0, 0)',
                    layer='below'
                )
            ],
            height=350,
            title={
                'y': 0.95,
                'x': 0.5,
                'xanchor': 'center',
                'yanchor': 'top',
                'font': dict(size=16, color="#1D1D1F")
            },
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
                font=dict(size=11)
            )
        )

        fig_purchase_progress.update_xaxes(
            title_text="月份",
            gridcolor='rgba(0, 0, 0, 0.05)',
            showline=False,
            linewidth=1,
            linecolor='rgba(0, 0, 0, 0.1)',
            tickfont=dict(size=11)
        )

        fig_purchase_progress.update_yaxes(
            title_text="采购业绩完成进度",
            gridcolor='rgba(0, 0, 0, 0.05)',
            showline=False,
            linewidth=1,
            linecolor='rgba(0, 0, 0, 0.1)',
            tickformat=".0%",
            tickfont=dict(size=11)
        )

        # 显示采购业绩完成进度折线图
        st.plotly_chart(fig_purchase_progress, use_container_width=True)


# 显示历月拿货量对比页面
def show_history_delivery_compare():
    title_card("历月拿货量对比分析")

    # 检查是否有历月数据
    if not st.session_state.get('history_files', []) or not st.session_state.get('history_data', {}):
        show_warning_card("请先在历月对比页面上传历月数据文件")
        return

    # 第一部分：月度拿货量对比
    st.markdown("""
    <h2 style="margin: 20px 0 15px 0; padding: 0; font-size: 28px; font-weight: 700; color: #1D1D1F; font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;">
        一、📊 月度拿货量对比
    </h2>
    """, unsafe_allow_html=True)

    # 获取所有上传的文件数据
    history_files = st.session_state['history_files']
    history_data = st.session_state['history_data']

    # 提取每个文件的有效拿货量数据
    delivery_data = []
    for file_info in history_files:
        file_name = file_info['name']
        file_data = history_data.get(file_name)

        if file_data and '本月拿货统计' in file_data:
            stats_df = file_data['本月拿货统计']

            # 获取姓名列
            name_col = '姓名' if '姓名' in stats_df.columns else ('报价员' if '报价员' in stats_df.columns else None)

            if name_col and '合计' in stats_df[name_col].values and '有效拿货量' in stats_df.columns:
                # 提取合计行数据
                total_row = stats_df[stats_df[name_col] == '合计'].copy()

                if not total_row.empty:
                    try:
                        # 提取有效拿货量
                        delivery_value = total_row['有效拿货量'].values[0]

                        # 处理可能的字符串格式
                        if isinstance(delivery_value, str):
                            delivery_value = float(delivery_value.replace(',', '').replace('¥', ''))

                        # 尝试从文件名提取年月信息
                        import re
                        match = re.search(r'采购(\d{4})年_(\d{1,2})月统计表', file_name)
                        if match:
                            year = match.group(1)
                            month = match.group(2)
                            month_label = f"{year}年{month}月"
                        else:
                            # 如果无法从文件名提取，则使用文件名作为标签
                            month_label = file_name.replace("采购", "").replace("统计表.xlsx", "")

                        delivery_data.append({
                            'month_label': month_label,
                            'file_name': file_name,
                            'delivery_amount': float(delivery_value)
                        })
                    except Exception as e:
                        st.warning(f"处理文件 '{file_name}' 时出错: {str(e)}")

    # 如果没有提取到数据
    if not delivery_data:
        show_warning_card(
            "未能从上传的文件中提取到有效的拿货量数据，请确保Excel文件包含'本月拿货统计'工作表和'有效拿货量'列")
        return

    # 按月份排序（假设月份标签格式为"YYYY年MM月"）
    try:
        delivery_data.sort(key=lambda x: (
            int(x['month_label'].split('年')[0]) if '年' in x['month_label'] else 0,
            int(x['month_label'].split('年')[1].split('月')[0]) if '年' in x['month_label'] and '月' in x[
                'month_label'] else 0
        ))
    except:
        # 如果排序失败，保持原顺序
        pass

    # 提取排序后的月份标签和拿货量数据
    months = [item['month_label'] for item in delivery_data]
    delivery_amounts = [item['delivery_amount'] for item in delivery_data]

    # 使用Plotly创建交互式折线图
    import plotly.graph_objects as go

    # 创建拿货量折线图
    fig_delivery = go.Figure()

    # 添加拿货量折线图
    fig_delivery.add_trace(
        go.Scatter(
            x=months,
            y=delivery_amounts,
            name="有效拿货量",
            line=dict(color="#30D158", width=3),  # 使用绿色
            mode='lines+markers',
            marker=dict(size=10, symbol='circle', color="#30D158",
                        line=dict(color='white', width=2))
        )
    )

    # 添加图表样式
    fig_delivery.update_layout(
        title_text="历月有效拿货量趋势分析",
        font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif", size=14),
        plot_bgcolor='rgba(255, 255, 255, 0.8)',
        paper_bgcolor='white',
        hovermode="x unified",
        margin=dict(l=30, r=30, t=100, b=30),
        shapes=[
            # 外框
            dict(
                type='rect',
                xref='paper', yref='paper',
                x0=-0.05, y0=-0.05, x1=1.05, y1=1.05,
                line=dict(color='rgba(0, 0, 0, 0.08)', width=1),
                fillcolor='rgba(0, 0, 0, 0)',
                layer='below'
            )
        ],
        width=None,
        height=650,
        title={
            'y': 0.95,
            'x': 0.5,
            'xanchor': 'center',
            'yanchor': 'top',
            'font': dict(size=18, color="#1D1D1F")
        },
    )

    # 更新轴标签和样式
    fig_delivery.update_xaxes(
        title_text="月份",
        gridcolor='rgba(0, 0, 0, 0.05)',
        showline=False,
        linewidth=1,
        linecolor='rgba(0, 0, 0, 0.1)',
        tickfont=dict(family="SF Pro Display", size=12)
    )

    fig_delivery.update_yaxes(
        title_text="拿货量",
        gridcolor='rgba(0, 0, 0, 0.05)',
        showline=False,
        linewidth=1,
        linecolor='rgba(0, 0, 0, 0.1)',
        tickformat=",.0f"
    )

    # 显示折线图
    st.plotly_chart(fig_delivery, use_container_width=True)

    # 创建表格数据
    delivery_df = pd.DataFrame({
        '月份': months,
        '有效拿货量': [f"{amount:,.0f}" for amount in delivery_amounts]
    })

    # 显示拿货量表格
    st.markdown("<h4 style='text-align: center; margin: 10px 0;'>历月有效拿货量明细</h4>",
                unsafe_allow_html=True)
    st.dataframe(delivery_df, use_container_width=True, hide_index=True)

    # 计算环比增长率
    if len(delivery_amounts) > 1:
        growth_rates = []
        for i in range(1, len(delivery_amounts)):
            if delivery_amounts[i - 1] != 0:
                growth_rate = (delivery_amounts[i] - delivery_amounts[i - 1]) / delivery_amounts[i - 1] * 100
                # 添加正号(+)显示正数增长率
                if growth_rate > 0:
                    growth_rates.append(f"+{growth_rate:.2f}%")
                else:
                    growth_rates.append(f"{growth_rate:.2f}%")
            else:
                growth_rates.append("无可用上月数据")

        # 为第一个月添加"无可用上月数据"
        growth_rates = ["无可用上月数据"] + growth_rates

        # 添加环比增长率到表格
        delivery_df['环比增长率'] = growth_rates

        # 显示带环比增长率的表格
        st.markdown("<h4 style='text-align: center; margin: 20px 0 10px 0;'>历月有效拿货量环比分析</h4>",
                    unsafe_allow_html=True)

        # 为环比增长率列应用样式，灰色显示"无可用上月数据"
        def highlight_unavailable(val):
            if val == "无可用上月数据":
                return 'color: gray'
            elif "+" in val:
                return 'color: green'
            elif "-" in val:
                return 'color: red'
            return ''

        # 应用样式并显示
        styled_df = delivery_df.style.map(highlight_unavailable, subset=['环比增长率'])
        st.dataframe(styled_df, use_container_width=True, hide_index=True)

    # 第二部分：月度拿货量分析
    st.markdown("""
    <h2 style="margin: 40px 0 15px 0; padding: 0; font-size: 28px; font-weight: 700; color: #1D1D1F; font-family: 'SF Pro Display', -apple-system, BlinkMacSystemFont, sans-serif;">
        二、📈 月度拿货量分析
    </h2>
    """, unsafe_allow_html=True)

    # 获取所有上传的文件数据
    history_files = st.session_state['history_files']
    history_data = st.session_state['history_data']

    # 提取所有员工名称
    all_employees = set()
    employee_data = {}

    # 收集各月份各员工的拿货数据
    for file_info in history_files:
        file_name = file_info['name']
        file_data = history_data.get(file_name)

        # 尝试从文件名提取年月信息
        match = re.search(r'采购(\d{4})年_(\d{1,2})月统计表', file_name)
        if match:
            year = match.group(1)
            month = match.group(2)
            month_label = f"{year}年{month}月"
        else:
            # 如果无法从文件名提取，则使用文件名作为标签
            month_label = file_name.replace("采购", "").replace("统计表.xlsx", "")

        if file_data and '本月拿货统计' in file_data:
            stats_df = file_data['本月拿货统计']

            # 获取姓名列
            name_col = '姓名' if '姓名' in stats_df.columns else ('报价员' if '报价员' in stats_df.columns else None)

            if name_col and '有效拿货量' in stats_df.columns and '目标拿货量' in stats_df.columns:
                # 过滤掉合计行
                employee_rows = stats_df[stats_df[name_col] != '合计'].copy()

                for _, row in employee_rows.iterrows():
                    employee_name = row[name_col]
                    if pd.notna(employee_name) and employee_name.strip():
                        all_employees.add(employee_name)

                        # 提取有效拿货量和目标拿货量
                        delivery_amount = row['有效拿货量'] if pd.notna(row['有效拿货量']) else 0
                        target_amount = row['目标拿货量'] if pd.notna(row['目标拿货量']) else 0

                        # 计算目标完成率
                        completion_rate = 0
                        if target_amount > 0:
                            completion_rate = (delivery_amount / target_amount) * 100

                        # 存储数据
                        if employee_name not in employee_data:
                            employee_data[employee_name] = []

                        employee_data[employee_name].append({
                            'month_label': month_label,
                            'delivery_amount': float(delivery_amount),
                            'target_amount': float(target_amount),
                            'completion_rate': completion_rate
                        })

    # 如果没有提取到数据
    if not all_employees:
        show_warning_card("未能从上传的文件中提取到有效的员工拿货量数据")
        return

    # 将员工名称列表排序
    sorted_employees = sorted(list(all_employees))

    # 添加多选框，允许用户选择多个员工
    st.subheader("选择员工查看拿货量分析")
    selected_employees = st.multiselect(
        "选择一个或多个员工",
        options=sorted_employees,
        default=[sorted_employees[0]] if sorted_employees else [],
        key="delivery_employee_selector"
    )

    # 如果没有选择员工，显示提示
    if not selected_employees:
        st.info("请选择至少一名员工查看拿货量分析")
        return

    # 创建左右分栏
    col1, col2 = st.columns(2)

    # 准备图表数据
    with col1:
        st.subheader("历月有效拿货量")

        # 创建有效拿货量折线图
        import plotly.graph_objects as go
        fig_delivery = go.Figure()

        # 为每个选定的员工添加折线
        color_palette = ["#0A84FF", "#30D158", "#FF453A", "#BF5AF2", "#FFD60A", "#66D4CF", "#FF9F0A", "#5E5CE6"]

        for i, employee in enumerate(selected_employees):
            if employee in employee_data:
                # 获取该员工的数据
                employee_entries = employee_data[employee]

                # 按月份排序
                try:
                    employee_entries.sort(key=lambda x: (
                        int(x['month_label'].split('年')[0]) if '年' in x['month_label'] else 0,
                        int(x['month_label'].split('年')[1].split('月')[0]) if '年' in x['month_label'] and '月' in x[
                            'month_label'] else 0
                    ))
                except:
                    pass  # 如果排序失败，保持原顺序

                # 提取月份和拿货量数据
                months = [item['month_label'] for item in employee_entries]
                delivery_amounts = [item['delivery_amount'] for item in employee_entries]

                # 添加到折线图
                color_index = i % len(color_palette)
                fig_delivery.add_trace(
                    go.Scatter(
                        x=months,
                        y=delivery_amounts,
                        name=employee,
                        line=dict(color=color_palette[color_index], width=2),
                        mode='lines+markers',
                        marker=dict(size=8, symbol='circle', color=color_palette[color_index],
                                    line=dict(color='white', width=1))
                    )
                )

        # 添加图表样式
        fig_delivery.update_layout(
            font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif", size=12),
            plot_bgcolor='rgba(255, 255, 255, 0.8)',
            paper_bgcolor='white',
            hovermode="x unified",
            margin=dict(l=20, r=20, t=50, b=30),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
                font=dict(size=10)
            ),
            height=400
        )

        # 更新轴标签和样式
        fig_delivery.update_xaxes(
            title_text="月份",
            gridcolor='rgba(0, 0, 0, 0.05)',
            showline=False,
            linewidth=1,
            linecolor='rgba(0, 0, 0, 0.1)',
            tickfont=dict(size=10)
        )

        fig_delivery.update_yaxes(
            title_text="有效拿货量",
            gridcolor='rgba(0, 0, 0, 0.05)',
            showline=False,
            linewidth=1,
            linecolor='rgba(0, 0, 0, 0.1)',
            tickformat=",.0f",
            tickfont=dict(size=10)
        )

        # 显示折线图
        st.plotly_chart(fig_delivery, use_container_width=True)

    with col2:
        st.subheader("拿货目标完成率")

        # 创建拿货目标完成率折线图
        fig_completion = go.Figure()

        # 为每个选定的员工添加折线
        for i, employee in enumerate(selected_employees):
            if employee in employee_data:
                # 获取该员工的数据
                employee_entries = employee_data[employee]

                # 按月份排序
                try:
                    employee_entries.sort(key=lambda x: (
                        int(x['month_label'].split('年')[0]) if '年' in x['month_label'] else 0,
                        int(x['month_label'].split('年')[1].split('月')[0]) if '年' in x['month_label'] and '月' in x[
                            'month_label'] else 0
                    ))
                except:
                    pass  # 如果排序失败，保持原顺序

                # 提取月份和完成率数据
                months = [item['month_label'] for item in employee_entries]
                completion_rates = [item['completion_rate'] for item in employee_entries]

                # 添加到折线图
                color_index = i % len(color_palette)
                fig_completion.add_trace(
                    go.Scatter(
                        x=months,
                        y=completion_rates,
                        name=employee,
                        line=dict(color=color_palette[color_index], width=2),
                        mode='lines+markers',
                        marker=dict(size=8, symbol='circle', color=color_palette[color_index],
                                    line=dict(color='white', width=1))
                    )
                )

        # 添加100%参考线
        fig_completion.add_shape(
            type="line",
            x0=0,
            y0=100,
            x1=1,
            y1=100,
            xref="paper",
            line=dict(
                color="rgba(0,0,0,0.3)",
                width=1,
                dash="dash",
            )
        )

        # 添加图表样式
        fig_completion.update_layout(
            font=dict(family="SF Pro Display, -apple-system, BlinkMacSystemFont, sans-serif", size=12),
            plot_bgcolor='rgba(255, 255, 255, 0.8)',
            paper_bgcolor='white',
            hovermode="x unified",
            margin=dict(l=20, r=20, t=50, b=30),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
                font=dict(size=10)
            ),
            height=400
        )

        # 更新轴标签和样式
        fig_completion.update_xaxes(
            title_text="月份",
            gridcolor='rgba(0, 0, 0, 0.05)',
            showline=False,
            linewidth=1,
            linecolor='rgba(0, 0, 0, 0.1)',
            tickfont=dict(size=10)
        )

        fig_completion.update_yaxes(
            title_text="拿货目标完成率(%)",
            gridcolor='rgba(0, 0, 0, 0.05)',
            showline=False,
            linewidth=1,
            linecolor='rgba(0, 0, 0, 0.1)',
            tickformat=".1f",
            tickfont=dict(size=10),
            range=[0, max(120, max([max([item['completion_rate'] for item in employee_data[emp]]) for emp in
                                    selected_employees if emp in employee_data]) * 1.1)]
        )

        # 显示折线图
        st.plotly_chart(fig_completion, use_container_width=True)

    # 显示详细数据表格
    st.subheader("详细数据")

    # 获取所有月份并排序
    all_months = set()
    for emp in employee_data:
        for entry in employee_data[emp]:
            all_months.add(entry['month_label'])

    sorted_months = sorted(list(all_months), key=lambda x: (
        int(x.split('年')[0]) if '年' in x else 0,
        int(x.split('年')[1].split('月')[0]) if '年' in x and '月' in x else 0
    ))

    # 准备表格数据：员工为行，月份为列
    table_data = {}
    for employee in selected_employees:
        if employee in employee_data:
            table_data[employee] = {}

            # 按月份整理该员工的数据
            for month in sorted_months:
                table_data[employee][month] = {
                    '有效拿货量': 0,
                    '目标完成率': 0
                }

            # 填充实际数据
            for entry in employee_data[employee]:
                month = entry['month_label']
                if month in table_data[employee]:
                    table_data[employee][month]['有效拿货量'] = entry['delivery_amount']
                    table_data[employee][month]['目标完成率'] = entry['completion_rate']

    # 转换为DataFrame格式
    rows = []
    for employee in selected_employees:
        if employee in table_data:
            row_data = {'姓名': employee}

            for month in sorted_months:
                row_data[f"{month}有效拿货量"] = f"{table_data[employee][month]['有效拿货量']:,.0f}"
                row_data[f"{month}目标完成率"] = f"{table_data[employee][month]['目标完成率']:.1f}%"

            rows.append(row_data)

    # 创建DataFrame并显示
    if rows:
        table_df = pd.DataFrame(rows)

        # 应用条件格式化：目标完成率超过100%显示为绿色，低于100%显示为红色
        def highlight_completion_rate(val):
            try:
                if '%' in str(val):
                    rate = float(str(val).replace('%', ''))
                    if rate >= 100:
                        return 'color: green; font-weight: bold'
                    else:
                        return 'color: red'
            except:
                pass
            return ''

        # 构建需要应用样式的列名列表（所有包含"目标完成率"的列）
        completion_rate_cols = [col for col in table_df.columns if '目标完成率' in col]

        # 应用样式
        styled_table = table_df.style.map(highlight_completion_rate, subset=completion_rate_cols)
        st.dataframe(styled_table, use_container_width=True, hide_index=True)


# 保存当前状态到撤销历史记录
def save_current_state():
    """
    保存当前会话状态到撤销历史中
    保存关键状态变量，用于撤销功能
    """
    current_state = {
        'current_page': st.session_state.get('current_page', 'home'),
        'history_compare_subpage': st.session_state.get('history_compare_subpage', None),
        'history_files': st.session_state.get('history_files', []),
        'history_data': st.session_state.get('history_data', {}),
        'processed_history_files': st.session_state.get('processed_history_files', set()),
        'excel_data': st.session_state.get('excel_data', None),
        'data_loaded': st.session_state.get('data_loaded', False),
        'file_path': st.session_state.get('file_path', None)
    }

    # 添加到历史记录
    st.session_state['previous_state'].append(current_state)

    # 限制历史记录长度，防止内存占用过多
    if len(st.session_state['previous_state']) > 10:
        st.session_state['previous_state'] = st.session_state['previous_state'][-10:]


# 主函数
def main():
    # 加载CSS样式
    load_css()

    # 初始化会话状态变量
    if 'previous_state' not in st.session_state:
        st.session_state['previous_state'] = []

    # 显示导航栏（主页不显示）
    if st.session_state['current_page'] != 'home':
        show_navigation()

    # 根据当前页面状态显示不同页面
    if st.session_state['current_page'] == 'home':
        show_home_page()
    elif st.session_state['current_page'] == 'leaderboard':
        show_leaderboard_page()
    elif st.session_state['current_page'] == 'purchase_detail':
        show_purchase_detail_page()
    elif st.session_state['current_page'] == 'purchase_stats':
        show_purchase_stats_page()
    elif st.session_state['current_page'] == 'delivery_stats':
        show_delivery_stats_page()
    elif st.session_state['current_page'] == 'history_compare':
        # 检查是否有子页面
        if st.session_state.get('history_compare_subpage') == 'history_purchase_compare':
            show_history_purchase_compare()
        elif st.session_state.get('history_compare_subpage') == 'history_delivery_compare':
            show_history_delivery_compare()
        else:
            show_history_compare_page()


if __name__ == '__main__':
    main()
