import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="CSV Cleaner",
    page_icon="🧹",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    * {
        font-family: 'Inter', sans-serif;
    }
    
    .main {
        background: linear-gradient(135deg, #f5f7fa 0%, #e4e8ec 100%);
        min-height: 100vh;
        padding: 2rem;
    }
    
    .container {
        max-width: 680px;
        margin: 0 auto;
    }
    
    .card {
        background: white;
        border-radius: 16px;
        padding: 2.5rem;
        box-shadow: 0 4px 24px rgba(0, 0, 0, 0.08);
    }
    
    .title {
        font-size: 1.75rem;
        font-weight: 700;
        color: #1a1a2e;
        margin-bottom: 0.5rem;
        text-align: center;
    }
    
    .subtitle {
        color: #6b7280;
        text-align: center;
        margin-bottom: 2rem;
        font-size: 0.95rem;
    }
    
    .upload-box {
        border: 2px dashed #d1d5db;
        border-radius: 12px;
        padding: 2rem;
        text-align: center;
        transition: all 0.3s ease;
        background: #fafafa;
        margin-bottom: 1.5rem;
    }
    
    .upload-box:hover {
        border-color: #6366f1;
        background: #f5f3ff;
    }
    
    .stFileUploader > div > div > div {
        background: transparent !important;
    }
    
    .stFileUploader label {
        display: none !important;
    }
    
    .btn-primary {
        background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%);
        color: white;
        border: none;
        padding: 0.875rem 2rem;
        border-radius: 10px;
        font-weight: 600;
        font-size: 1rem;
        cursor: pointer;
        transition: all 0.3s ease;
        width: 100%;
    }
    
    .btn-primary:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 24px rgba(99, 102, 241, 0.35);
    }
    
    .btn-primary:disabled {
        opacity: 0.5;
        cursor: not-allowed;
        transform: none;
        box-shadow: none;
    }
    
    .select-box {
        margin-bottom: 1.5rem;
    }
    
    .select-box label {
        display: block;
        font-weight: 500;
        color: #374151;
        margin-bottom: 0.5rem;
        font-size: 0.9rem;
    }
    
    .stSelectbox > div > div {
        border-radius: 8px !important;
        border-color: #d1d5db !important;
    }
    
    .stats {
        display: grid;
        grid-template-columns: 1fr 1fr 1fr;
        gap: 1rem;
        margin: 1.5rem 0;
    }
    
    .stat-box {
        background: #f9fafb;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
    }
    
    .stat-value {
        font-size: 1.5rem;
        font-weight: 700;
        color: #6366f1;
    }
    
    .stat-label {
        font-size: 0.75rem;
        color: #6b7280;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .success-box {
        background: #ecfdf5;
        border: 1px solid #a7f3d0;
        border-radius: 10px;
        padding: 1rem;
        margin: 1.5rem 0;
        color: #065f46;
    }
    
    .download-btn {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white;
        border: none;
        padding: 0.875rem 2rem;
        border-radius: 10px;
        font-weight: 600;
        font-size: 1rem;
        cursor: pointer;
        transition: all 0.3s ease;
        display: block;
        width: 100%;
        text-align: center;
        text-decoration: none;
    }
    
    .download-btn:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 24px rgba(16, 185, 129, 0.35);
    }
    
    .divider {
        height: 1px;
        background: #e5e7eb;
        margin: 2rem 0;
    }
    
    .info-text {
        font-size: 0.85rem;
        color: #6b7280;
        line-height: 1.6;
    }
</style>
""", unsafe_allow_html=True)

def has_address(value):
    if pd.isna(value) or str(value).strip() == "":
        return False
    val = str(value).strip().lower()
    address_indicators = [
        'street', 'st,', 'st.', 'avenue', 'ave', 'road', 'rd', 'boulevard', 'blvd',
        'drive', 'dr', 'lane', 'ln', 'court', 'ct', 'place', 'pl', 'way',
        'floor', 'suite', 'ste', 'unit', 'building', 'bldg',
        'apartment', 'apt', 'house', 'box', 'po box',
        'city', 'state', 'zip', 'postal',
        'north', 'south', 'east', 'west', 'n ', ' s ', ' e ', ' w ',
        '#', 'unit ', 'apt ', 'ste '
    ]
    return any(indicator in val for indicator in address_indicators)

def is_action_item(value):
    if pd.isna(value) or str(value).strip() == "":
        return False
    val = str(value).strip().lower()
    action_indicators = [
        'call', 'email', 'send', 'follow up', 'followup', 'check', 'review',
        'submit', 'complete', 'finish', 'update', 'prepare', 'schedule',
        'arrange', 'confirm', 'verify', 'process', 'create', 'draft',
        'research', 'investigate', 'analyze', 'organize', 'file',
        'pick up', 'drop off', 'deliver', 'install', 'fix', 'repair',
        'clean', 'paint', 'replace', 'order', 'purchase', 'buy'
    ]
    return any(val.startswith(indicator) or indicator in val for indicator in action_indicators)

st.markdown('<div class="container"><div class="card">', unsafe_allow_html=True)
st.markdown('<div class="title">🧹 CSV Cleaner</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Clean your Asana exports by removing subtasks</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("", type=["csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file, dtype=str)
    
    columns = list(df.columns)
    
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        address_col = st.selectbox(
            "Address column",
            options=columns,
            index=0 if "address" in " ".join(columns).lower() else 0,
            help="Select the column containing addresses to identify primary tasks"
        )
    with col2:
        name_col = st.selectbox(
            "Task name column",
            options=columns,
            index=1 if len(columns) > 1 else 0,
            help="Select the task name column to detect action items"
        )
    
    st.markdown(f'<p class="info-text">Primary tasks are identified by having an address in "{address_col}". Subtasks (action items) without addresses will be removed.</p>', unsafe_allow_html=True)
    
    if st.button("Clean CSV", disabled=False):
        original_count = len(df)
        
        mask = df[address_col].apply(has_address)
        filtered_df = df[mask].copy()
        
        final_count = len(filtered_df)
        removed_count = original_count - final_count
        
        if 'Parent task' in filtered_df.columns:
            cols = list(filtered_df.columns)
            parent_idx = cols.index('Parent task')
            cols.remove('Parent task')
            cols.insert(1, 'Parent task')
            filtered_df = filtered_df[cols]
        
        buffer = io.StringIO()
        filtered_df.to_csv(buffer, index=False)
        buffer.seek(0)
        
        st.markdown(f"""
        <div class="stats">
            <div class="stat-box">
                <div class="stat-value">{original_count}</div>
                <div class="stat-label">Original</div>
            </div>
            <div class="stat-box">
                <div class="stat-value">{final_count}</div>
                <div class="stat-label">Cleaned</div>
            </div>
            <div class="stat-box">
                <div class="stat-value">{removed_count}</div>
                <div class="stat-label">Removed</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown(f'<div class="success-box">✅ Cleaned file ready with {final_count} primary tasks</div>', unsafe_allow_html=True)
        
        st.download_button(
            label="Download Cleaned CSV",
            data=buffer.getvalue(),
            file_name="cleaned_" + uploaded_file.name,
            mime="text/csv"
        )

st.markdown('</div></div>', unsafe_allow_html=True)
