"""
================================================================
  Triune Entertainment – Budget Analysis Tool (Web Version v2)
  IT493 | Team 4
================================================================
  FIXED:
  - Proper Budget (left col 7) vs Actual (right col 13) extraction
  - All 6 professional charts with enhanced quality
  - Password protection (default: triune2024)
  - Multiple file upload support (with unique keys)
  - Smart delta colors (red=bad, green=good)
================================================================
"""

import streamlit as st
import os, io, json, hashlib
from datetime import datetime
import pandas as pd

# Configure matplotlib BEFORE importing pyplot
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

# ── Brand Colors ──────────────────────────────────────────────
NAVY="#1F3864"; TEAL="#2E75B6"; LIGHT="#D6E4F0"; WHITE="#FFFFFF"
GREEN="#70AD47"; RED="#C00000"; DARK="#152848"; PURPLE="#9B59B6"
GOLD="#FFF2CC"

OX_NAVY="1F3864"; OX_TEAL="2E75B6"; OX_LIGHT="D6E4F0"
OX_WHITE="FFFFFF"; OX_GOLD="FFF2CC"

# ── Page Configuration ────────────────────────────────────────
st.set_page_config(
    page_title="Triune Budget Analysis Tool",
    page_icon="🎭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Custom CSS ─────────────────────────────────────────────────
st.markdown("""
    <style>
    .main {background-color: #F0F5FB;}
    .stButton>button {
        background-color: #2E75B6; color: white; font-weight: bold;
        border-radius: 5px; padding: 10px 20px; border: none;
    }
    .stButton>button:hover {background-color: #1F3864;}
    h1 {color: #1F3864; font-family: Georgia, serif;}
    h2 {color: #2E75B6; font-family: Georgia, serif;}
    </style>
    """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════
#  PASSWORD PROTECTION
# ═══════════════════════════════════════════════════

def check_password():
    """Returns True if user entered correct password."""
    
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
    
    if st.session_state["password_correct"]:
        return True
    
    st.markdown("""
        <div style='background-color: #1F3864; padding: 30px; border-radius: 10px; margin-bottom: 20px; text-align: center;'>
            <h1 style='color: white; margin: 0;'>🎭 Triune Entertainment</h1>
            <h2 style='color: #D6E4F0; margin: 10px 0;'>Budget Analysis Tool</h2>
            <p style='color: #2E75B6; margin: 0; font-family: monospace;'>IT493 | Team 4</p>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 🔒 Secure Login")
    st.markdown("Please enter your password to access the Budget Analysis Tool.")
    
    password = st.text_input("Password", type="password", key="password_input")
    
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        if st.button("🔓 Login", use_container_width=True):
            if password == "triune2024":
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("❌ Incorrect password. Please try again.")
    
    st.markdown("---")
    st.info("Contact your IT administrator for login credentials.")
    
    return False


# ═══════════════════════════════════════════════════
#  DATA EXTRACTION
# ═══════════════════════════════════════════════════

def extract_budget_data(uploaded_file):
    """Extract Budget (LEFT col 7) and Actual (RIGHT col 13) data."""
    try:
        df = pd.read_excel(uploaded_file, header=None)
    except Exception as e:
        return None, f"Cannot read file: {e}"
    
    # Extract show name
    show_name = "Unknown Show"
    show_date = ""
    
    for col in range(df.shape[1]):
        val = df.iloc[1, col]
        if pd.notna(val) and isinstance(val, str) and len(val) > 5:
            show_name = val.strip()
            if " - " in show_name:
                parts = show_name.split(" - ")
                show_name = parts[0].strip()
                show_date = parts[1].strip() if len(parts) > 1 else ""
            if "Director:" in show_name:
                show_name = show_name.split("Director:")[0].strip()
            break
    
    # Extract revenue: LEFT (col 7) = BUDGET, RIGHT (col 13) = ACTUAL
    budget_revenue = actual_revenue = 0
    for idx in range(len(df)):
        row_text = str(df.iloc[idx, 1]) if pd.notna(df.iloc[idx, 1]) else ""
        if "Total 4300 Revenues" in row_text:
            budget_revenue = pd.to_numeric(df.iloc[idx, 7], errors='coerce') or 0
            actual_revenue = pd.to_numeric(df.iloc[idx, 13], errors='coerce') or 0
            break
    
    # Extract expenses: LEFT (col 7) = BUDGET, RIGHT (col 13) = ACTUAL
    budget_expenses = actual_expenses = 0
    for idx in range(len(df)):
        row_text = str(df.iloc[idx, 1]) if pd.notna(df.iloc[idx, 1]) else ""
        if "Total 5000 Direct Production Costs" in row_text:
            budget_expenses = pd.to_numeric(df.iloc[idx, 7], errors='coerce') or 0
            actual_expenses = pd.to_numeric(df.iloc[idx, 13], errors='coerce') or 0
            break
    
    # Calculate metrics
    budget_net = budget_revenue - budget_expenses
    actual_net = actual_revenue - actual_expenses
    budget_margin = (budget_net / budget_revenue * 100) if budget_revenue > 0 else 0
    actual_margin = (actual_net / actual_revenue * 100) if actual_revenue > 0 else 0
    revenue_variance = actual_revenue - budget_revenue
    expense_variance = actual_expenses - budget_expenses
    net_variance = actual_net - budget_net
    
    return {
        'filename': uploaded_file.name,
        'show_name': show_name,
        'show_date': show_date,
        'budget_revenue': budget_revenue,
        'actual_revenue': actual_revenue,
        'budget_expenses': budget_expenses,
        'actual_expenses': actual_expenses,
        'budget_net': budget_net,
        'actual_net': actual_net,
        'budget_margin': budget_margin,
        'actual_margin': actual_margin,
        'revenue_variance': revenue_variance,
        'expense_variance': expense_variance,
        'net_variance': net_variance,
        'revenue_variance_pct': (revenue_variance / budget_revenue * 100) if budget_revenue > 0 else 0,
        'expense_variance_pct': (expense_variance / budget_expenses * 100) if budget_expenses > 0 else 0,
    }, None


# ═══════════════════════════════════════════════════
#  CHART FUNCTIONS (6 total)
# ═══════════════════════════════════════════════════

def create_chart_1_budget_vs_actual(data):
    """Chart 1: Budget vs Actual - 3 comparisons"""
    fig, (ax1, ax2, ax3) = plt.subplots(1, 3, figsize=(18, 6))
    fig.patch.set_facecolor('#F0F5FB')
    
    x = [0, 1]
    
    # Revenue
    ax1.set_facecolor('#FFFFFF')
    bars1 = ax1.bar(x, [data['budget_revenue'], data['actual_revenue']],
                    color=[TEAL, PURPLE], alpha=0.9, width=0.6, linewidth=2, edgecolor='white')
    ax1.set_title('Revenue: Budget vs Actual', fontsize=14, fontweight='bold', color=NAVY, pad=15)
    ax1.set_xticks(x)
    ax1.set_xticklabels(['Budget', 'Actual'])
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'${v:,.0f}'))
    ax1.grid(axis='y', alpha=0.3, linestyle='--')
    for bar in bars1:
        h = bar.get_height()
        ax1.annotate(f'${h:,.0f}', xy=(bar.get_x()+bar.get_width()/2, h),
                    xytext=(0,5), textcoords='offset points', ha='center', fontsize=10, fontweight='bold')
    
    # Expenses
    ax2.set_facecolor('#FFFFFF')
    bars2 = ax2.bar(x, [data['budget_expenses'], data['actual_expenses']],
                    color=[NAVY, RED], alpha=0.9, width=0.6, linewidth=2, edgecolor='white')
    ax2.set_title('Expenses: Budget vs Actual', fontsize=14, fontweight='bold', color=NAVY, pad=15)
    ax2.set_xticks(x)
    ax2.set_xticklabels(['Budget', 'Actual'])
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'${v:,.0f}'))
    ax2.grid(axis='y', alpha=0.3, linestyle='--')
    for bar in bars2:
        h = bar.get_height()
        ax2.annotate(f'${h:,.0f}', xy=(bar.get_x()+bar.get_width()/2, h),
                    xytext=(0,5), textcoords='offset points', ha='center', fontsize=10, fontweight='bold')
    
    # Net Income
    ax3.set_facecolor('#FFFFFF')
    colors = [TEAL if v>=0 else RED for v in [data['budget_net'], data['actual_net']]]
    bars3 = ax3.bar(x, [data['budget_net'], data['actual_net']],
                    color=colors, alpha=0.9, width=0.6, linewidth=2, edgecolor='white')
    ax3.set_title('Net Income: Budget vs Actual', fontsize=14, fontweight='bold', color=NAVY, pad=15)
    ax3.set_xticks(x)
    ax3.set_xticklabels(['Budget', 'Actual'])
    ax3.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'${v:,.0f}'))
    ax3.axhline(y=0, color='black', linewidth=1)
    ax3.grid(axis='y', alpha=0.3, linestyle='--')
    for bar in bars3:
        h = bar.get_height()
        ax3.annotate(f'${h:,.0f}', xy=(bar.get_x()+bar.get_width()/2, h),
                    xytext=(0,5 if h>=0 else -15), textcoords='offset points', ha='center', fontsize=10, fontweight='bold')
    
    plt.tight_layout()
    return fig

def create_chart_2_variance(data):
    """Chart 2: Variance Analysis"""
    fig, ax = plt.subplots(figsize=(12, 7))
    fig.patch.set_facecolor('#F0F5FB')
    ax.set_facecolor('#FFFFFF')
    
    categories = ['Revenue', 'Expenses', 'Net Income']
    variances = [data['revenue_variance'], data['expense_variance'], data['net_variance']]
    colors = [GREEN if v>=0 else RED for v in variances]
    
    bars = ax.bar(categories, variances, color=colors, alpha=0.9, width=0.5, linewidth=2, edgecolor='white')
    ax.set_title('Variance Analysis (Actual - Budget)', fontsize=16, fontweight='bold', color=NAVY, pad=20)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'${v:,.0f}'))
    ax.axhline(y=0, color='black', linewidth=2)
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    
    for bar, val in zip(bars, variances):
        h = bar.get_height()
        pct = (val / data['budget_revenue'] * 100) if data['budget_revenue'] > 0 else 0
        ax.annotate(f'${abs(h):,.0f}\n({abs(pct):.1f}%)',
                   xy=(bar.get_x()+bar.get_width()/2, h),
                   xytext=(0, 10 if h>=0 else -35), textcoords='offset points',
                   ha='center', fontsize=11, fontweight='bold')
    
    plt.tight_layout()
    return fig

def create_chart_3_pie(data):
    """Chart 3: Pie Charts"""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))
    fig.patch.set_facecolor('#F0F5FB')
    
    # Budget Pie
    sizes1 = [data['budget_net'], data['budget_expenses']]
    labels1 = [f'Net ${sizes1[0]:,.0f}', f'Expenses ${sizes1[1]:,.0f}']
    wedges1, texts1, autotexts1 = ax1.pie(sizes1, labels=labels1, autopct='%1.1f%%',
            colors=[TEAL, NAVY], explode=(0.1,0), shadow=True, startangle=90)
    ax1.set_title(f'Budget\nTotal: ${data["budget_revenue"]:,.0f}', fontsize=14, fontweight='bold', pad=15)
    
    # Actual Pie
    sizes2 = [data['actual_net'], data['actual_expenses']]
    labels2 = [f'Net ${sizes2[0]:,.0f}', f'Expenses ${sizes2[1]:,.0f}']
    wedges2, texts2, autotexts2 = ax2.pie(sizes2, labels=labels2, autopct='%1.1f%%',
            colors=[PURPLE, RED], explode=(0.1,0), shadow=True, startangle=90)
    ax2.set_title(f'Actual\nTotal: ${data["actual_revenue"]:,.0f}', fontsize=14, fontweight='bold', pad=15)
    
    plt.tight_layout()
    return fig

def create_chart_4_scatter(data):
    """Chart 4: Scatter Plot"""
    fig, ax = plt.subplots(figsize=(11, 8))
    fig.patch.set_facecolor('#F0F5FB')
    ax.set_facecolor('#FFFFFF')
    
    categories = ['Revenue', 'Expenses', 'Net']
    budget = [data['budget_revenue'], data['budget_expenses'], data['budget_net']]
    actual = [data['actual_revenue'], data['actual_expenses'], data['actual_net']]
    colors = [TEAL, NAVY, PURPLE]
    
    for cat, b, a, c in zip(categories, budget, actual, colors):
        ax.scatter(b, a, s=300, c=c, alpha=0.7, edgecolors='white', linewidth=2, label=cat)
        ax.annotate(cat, (b,a), xytext=(10,10), textcoords='offset points', fontsize=9)
    
    max_val = max(budget + actual) * 1.1
    min_val = min(budget + actual) * 0.9
    ax.plot([min_val,max_val], [min_val,max_val], 'k--', alpha=0.5, linewidth=2, label='Perfect')
    
    ax.set_xlabel('Budget ($)', fontsize=12, fontweight='bold')
    ax.set_ylabel('Actual ($)', fontsize=12, fontweight='bold')
    ax.set_title('Budget Accuracy', fontsize=15, fontweight='bold', color=NAVY, pad=15)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'${v:,.0f}'))
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'${v:,.0f}'))
    ax.legend()
    ax.grid(True, alpha=0.3)
    
    plt.tight_layout()
    return fig

def create_chart_5_line(data):
    """Chart 5: Line Graph"""
    fig, ax = plt.subplots(figsize=(13, 7))
    fig.patch.set_facecolor('#F0F5FB')
    ax.set_facecolor('#FFFFFF')
    
    categories = ['Revenue', 'Expenses', 'Net Income']
    x_pos = [0,1,2]
    budget = [data['budget_revenue'], data['budget_expenses'], data['budget_net']]
    actual = [data['actual_revenue'], data['actual_expenses'], data['actual_net']]
    
    ax.plot(x_pos, budget, marker='o', markersize=12, linewidth=3, color=TEAL,
            label='Budget', markeredgecolor='white', markeredgewidth=2)
    ax.plot(x_pos, actual, marker='s', markersize=12, linewidth=3, color=PURPLE,
            label='Actual', markeredgecolor='white', markeredgewidth=2, linestyle='--')
    
    for i, (b,a) in enumerate(zip(budget, actual)):
        ax.annotate(f'${b:,.0f}', (i,b), xytext=(0,10), textcoords='offset points',
                   ha='center', fontsize=9, fontweight='bold', color=TEAL)
        ax.annotate(f'${a:,.0f}', (i,a), xytext=(0,-18), textcoords='offset points',
                   ha='center', fontsize=9, fontweight='bold', color=PURPLE)
    
    ax.set_xticks(x_pos)
    ax.set_xticklabels(categories, fontsize=11, fontweight='bold')
    ax.set_ylabel('Amount ($)', fontsize=12, fontweight='bold')
    ax.set_title('Trend Analysis', fontsize=15, fontweight='bold', color=NAVY, pad=15)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'${v:,.0f}'))
    ax.legend(fontsize=11)
    ax.grid(True, alpha=0.3, axis='y')
    
    plt.tight_layout()
    return fig

def create_chart_6_bar(data):
    """Chart 6: Comprehensive Bar Graph"""
    fig, ax = plt.subplots(figsize=(15, 8))
    fig.patch.set_facecolor('#F0F5FB')
    ax.set_facecolor('#FFFFFF')
    
    categories = ['Revenue', 'Expenses', 'Net Income']
    budget = [data['budget_revenue'], data['budget_expenses'], data['budget_net']]
    actual = [data['actual_revenue'], data['actual_expenses'], data['actual_net']]
    
    x = range(len(categories))
    width = 0.35
    
    bars1 = ax.bar([i-width/2 for i in x], budget, width, label='Budget', color=TEAL, alpha=0.9,
                   edgecolor='white', linewidth=2)
    bars2 = ax.bar([i+width/2 for i in x], actual, width, label='Actual', color=PURPLE, alpha=0.9,
                   edgecolor='white', linewidth=2)
    
    for bars in [bars1, bars2]:
        for bar in bars:
            h = bar.get_height()
            ax.annotate(f'${h:,.0f}', xy=(bar.get_x()+bar.get_width()/2, h),
                       xytext=(0, 5 if h>=0 else -15), textcoords='offset points',
                       ha='center', fontsize=10, fontweight='bold')
    
    # Variance labels
    for i, (b,a) in enumerate(zip(budget, actual)):
        if b != 0:
            var_pct = ((a-b)/b)*100
            color = GREEN if var_pct>=0 else RED
            symbol = "▲" if var_pct>=0 else "▼"
            ax.annotate(f'{symbol} {abs(var_pct):.1f}%', xy=(i, max(b,a)),
                       xytext=(0,25), textcoords='offset points', ha='center',
                       fontsize=12, fontweight='bold', color=color)
    
    ax.set_xlabel('Category', fontsize=13, fontweight='bold')
    ax.set_ylabel('Amount ($)', fontsize=13, fontweight='bold')
    ax.set_title('Complete Financial Comparison', fontsize=16, fontweight='bold', color=NAVY, pad=20)
    ax.set_xticks(x)
    ax.set_xticklabels(categories, fontsize=12, fontweight='bold')
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'${v:,.0f}'))
    ax.legend(fontsize=12)
    ax.grid(True, alpha=0.3, axis='y')
    ax.axhline(y=0, color='black', linewidth=1)
    
    plt.tight_layout()
    return fig


# ═══════════════════════════════════════════════════
#  EXCEL REPORT
# ═══════════════════════════════════════════════════

def generate_excel_report(data, charts_dict):
    """Generate Excel report with embedded charts."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    
    ws['A1'] = f"Show: {data['show_name']}"
    ws['A1'].font = Font(bold=True, size=16, color=OX_NAVY)
    ws['A2'] = f"Date: {data['show_date']}"
    ws['A3'] = f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}"
    
    ws['A5'] = "Category"
    ws['B5'] = "Budget"
    ws['C5'] = "Actual"
    ws['D5'] = "Variance"
    ws['E5'] = "Variance %"
    
    for col in ['A','B','C','D','E']:
        ws[f'{col}5'].fill = PatternFill('solid', fgColor=OX_TEAL)
        ws[f'{col}5'].font = Font(bold=True, color=OX_WHITE)
        ws[f'{col}5'].alignment = Alignment(horizontal='center')
    
    ws['A6'] = "Revenue"
    ws['B6'] = data['budget_revenue']
    ws['C6'] = data['actual_revenue']
    ws['D6'] = data['revenue_variance']
    ws['E6'] = data['revenue_variance_pct']/100
    
    ws['A7'] = "Expenses"
    ws['B7'] = data['budget_expenses']
    ws['C7'] = data['actual_expenses']
    ws['D7'] = data['expense_variance']
    ws['E7'] = data['expense_variance_pct']/100
    
    ws['A8'] = "Net Income"
    ws['B8'] = data['budget_net']
    ws['C8'] = data['actual_net']
    ws['D8'] = data['net_variance']
    ws['E8'] = (data['net_variance']/data['budget_net'])/100 if data['budget_net']!=0 else 0
    
    for row in range(6,9):
        for col in ['B','C','D']:
            ws[f'{col}{row}'].number_format = '$#,##0.00'
        ws[f'E{row}'].number_format = '0.00%'
        if row%2==0:
            for col in ['A','B','C','D','E']:
                ws[f'{col}{row}'].fill = PatternFill('solid', fgColor=OX_LIGHT)
    
    for col, width in [('A',15), ('B',18), ('C',18), ('D',18), ('E',15)]:
        ws.column_dimensions[col].width = width
    
    # Add each chart to separate sheets
    chart_names = ['Budget_vs_Actual', 'Variance_Analysis', 'Pie_Charts', 
                   'Scatter_Plot', 'Line_Graph', 'Bar_Graph']
    
    for chart_name, fig in zip(chart_names, charts_dict.values()):
        # Create new sheet
        ws_chart = wb.create_sheet(title=chart_name)
        
        # Save figure to bytes
        img_buffer = io.BytesIO()
        fig.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
        img_buffer.seek(0)
        
        # Add image to sheet
        img = XLImage(img_buffer)
        ws_chart.add_image(img, 'A1')
    
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ═══════════════════════════════════════════════════
#  MAIN APP
# ═══════════════════════════════════════════════════

def main():
    if not check_password():
        return
    
    st.markdown("""
        <div style='background-color: #1F3864; padding: 20px; border-radius: 10px; margin-bottom: 20px;'>
            <h1 style='color: white; text-align: center; margin: 0;'>🎭 Triune Entertainment</h1>
            <h2 style='color: #D6E4F0; text-align: center; margin: 5px 0 15px 0;'>Budget Analysis & Visualization Tool</h2>
            <div style='text-align: center;'>
                <a href='https://ads.google.com/aw/campaigns/new/performancemax?campaignId=281498546159968&ocid=8030665422' 
                   target='_blank' 
                   style='background-color: #2E75B6; color: white; padding: 10px 20px; margin: 0 10px; 
                          text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;'>
                    📢 Marketing Dashboard
                </a>
                <a href='https://airtable.com/appDuTE72UfHIFOJT/shrmRPv812yX1gDlZ/tblDl6p4SNVzHLQmZ/viw0wLiFVeOZYvJ6b' 
                   target='_blank' 
                   style='background-color: #70AD47; color: white; padding: 10px 20px; margin: 0 10px; 
                          text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;'>
                    👥 Talent CRM
                </a>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    with st.sidebar:
        st.markdown("### 📁 Upload Budget Worksheets")
        uploaded_files = st.file_uploader("Choose Excel files", type=['xlsx','xls'],
                                         accept_multiple_files=True)
        st.markdown("---")
        st.markdown("### 📊 Features")
        st.markdown("✅ 6 professional charts\n✅ Variance analysis\n✅ Multi-file support\n✅ Password protected")
        
        if st.button("🔓 Logout"):
            st.session_state["password_correct"] = False
            st.rerun()
    
    if not uploaded_files:
        st.info("👆 Upload budget worksheets to get started")
    else:
        st.markdown("### 📊 Analysis Results")
        
        for idx, uploaded_file in enumerate(uploaded_files):
            st.markdown("---")
            with st.expander(f"📄 {uploaded_file.name}", expanded=True):
                data, error = extract_budget_data(uploaded_file)
                
                if error:
                    st.error(f"❌ Error: {error}")
                    continue
                
                st.markdown(f"## {data['show_name']} ({data['show_date']})")
                
                col1, col2, col3, col4 = st.columns(4)
                
                # Pre-format all values
                budget_rev_str = f"${data['budget_revenue']:,.2f}"
                actual_rev_str = f"${data['actual_revenue']:,.2f}"
                rev_delta_str = f"${data['revenue_variance']:,.2f}"
                
                budget_exp_str = f"${data['budget_expenses']:,.2f}"
                actual_exp_str = f"${data['actual_expenses']:,.2f}"
                exp_delta_str = f"${data['expense_variance']:,.2f}"
                
                budget_net_str = f"${data['budget_net']:,.2f}"
                actual_net_str = f"${data['actual_net']:,.2f}"
                net_delta_str = f"${data['net_variance']:,.2f}"
                
                budget_margin_str = f"{data['budget_margin']:.1f}%"
                actual_margin_str = f"{data['actual_margin']:.1f}%"
                margin_diff = data['actual_margin'] - data['budget_margin']
                margin_delta_str = f"{margin_diff:.1f}%"
                
                with col1:
                    st.metric("Budget Revenue", budget_rev_str)
                    delta_color = "normal" if data['revenue_variance'] >= 0 else "inverse"
                    st.metric("Actual Revenue", actual_rev_str,
                             delta=rev_delta_str, delta_color=delta_color)
                
                with col2:
                    st.metric("Budget Expenses", budget_exp_str)
                    delta_color = "inverse" if data['expense_variance'] >= 0 else "normal"
                    st.metric("Actual Expenses", actual_exp_str,
                             delta=exp_delta_str, delta_color=delta_color)
                
                with col3:
                    st.metric("Budget Net", budget_net_str)
                    delta_color = "normal" if data['net_variance'] >= 0 else "inverse"
                    st.metric("Actual Net", actual_net_str,
                             delta=net_delta_str, delta_color=delta_color)
                
                with col4:
                    st.metric("Budget Margin", budget_margin_str)
                    delta_color = "normal" if margin_diff >= 0 else "inverse"
                    st.metric("Actual Margin", actual_margin_str,
                             delta=margin_delta_str, delta_color=delta_color)
                
                st.markdown("### 📈 All 6 Professional Charts")
                
                tab1,tab2,tab3,tab4,tab5,tab6 = st.tabs([
                    "Budget vs Actual", "Variance", "Pie Charts", 
                    "Scatter Plot", "Line Graph", "Bar Graph"
                ])
                
                # Create all charts and store them
                fig1 = create_chart_1_budget_vs_actual(data)
                fig2 = create_chart_2_variance(data)
                fig3 = create_chart_3_pie(data)
                fig4 = create_chart_4_scatter(data)
                fig5 = create_chart_5_line(data)
                fig6 = create_chart_6_bar(data)
                
                charts_dict = {
                    'budget_vs_actual': fig1,
                    'variance': fig2,
                    'pie': fig3,
                    'scatter': fig4,
                    'line': fig5,
                    'bar': fig6
                }
                
                with tab1:
                    st.pyplot(fig1)
                with tab2:
                    st.pyplot(fig2)
                with tab3:
                    st.pyplot(fig3)
                with tab4:
                    st.pyplot(fig4)
                with tab5:
                    st.pyplot(fig5)
                with tab6:
                    st.pyplot(fig6)
                
                # Close all figures after display
                for fig in charts_dict.values():
                    plt.close(fig)
                
                st.markdown("### 💾 Download Report")
                
                # Custom filename input
                default_filename = f"{data['show_name']}_{data['show_date']}_analysis"
                custom_filename = st.text_input(
                    "Report filename (without .xlsx extension):",
                    value=default_filename,
                    help="Enter a custom name for your report or use the default"
                )
                
                # Generate Excel with charts
                excel_data = generate_excel_report(data, charts_dict)
                st.download_button(
                    label=f"📥 Download {data['show_name']} Excel Report (with all 6 charts)",
                    data=excel_data,
                    file_name=f"{custom_filename}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

if __name__ == "__main__":
    main()
