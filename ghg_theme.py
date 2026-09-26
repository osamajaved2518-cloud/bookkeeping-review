"""Visual theme for the GHG Inventory dashboard.

Drop this file next to app.py. It adds: Inter font, card-style KPIs and charts,
a gradient hero banner, a branded sidebar header, finding cards and a footer.
"""
import html

import streamlit as st

GREEN = "#1E5631"
TEAL = "#2A9D8F"
LIME = "#A8C686"
GOLD = "#C9A227"
INK = "#1B1F2A"
GREY = "#6B7280"

AUTHOR = "Osama J."
AUTHOR_LINK = ""  # e.g. your Upwork or LinkedIn profile URL


def apply_theme():
    st.markdown(
        f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"], .stMarkdown, .stText, .stCaption, button, input, textarea, select {{
    font-family: 'Inter', 'Segoe UI', Arial, sans-serif !important;
}}

/* Hide Streamlit chrome */
#MainMenu, footer {{ visibility: hidden; }}
header[data-testid="stHeader"] {{ background: transparent; }}

/* Page */
.stApp {{ background: #F4F7F4; }}
.block-container {{ padding-top: 1.6rem; padding-bottom: 3rem; max-width: 1400px; }}
h1 {{ font-weight: 800 !important; color: {INK} !important; letter-spacing: -0.02em; }}
h4 {{ font-weight: 700 !important; color: {INK} !important; margin-top: 1.2rem !important; }}

/* Sidebar */
section[data-testid="stSidebar"] {{
    background: #FFFFFF;
    border-right: 1px solid #E5EAE5;
}}
section[data-testid="stSidebar"] .stRadio label p {{ font-weight: 500; }}

/* KPI cards */
div[data-testid="stMetric"] {{
    background: #FFFFFF;
    border: 1px solid #E5EAE5;
    border-left: 5px solid {GREEN};
    border-radius: 14px;
    padding: 16px 18px;
    box-shadow: 0 2px 10px rgba(27, 31, 42, 0.05);
    transition: transform .15s ease, box-shadow .15s ease;
}}
div[data-testid="stMetric"]:hover {{
    transform: translateY(-2px);
    box-shadow: 0 8px 22px rgba(27, 31, 42, 0.09);
}}
div[data-testid="stMetricLabel"] p {{
    font-size: 0.78rem !important; font-weight: 600 !important;
    text-transform: uppercase; letter-spacing: 0.06em; color: {GREY} !important;
}}
div[data-testid="stMetricValue"] {{ font-size: 1.9rem !important; font-weight: 800 !important; color: {INK}; }}

/* Chart and table cards */
div[data-testid="stPlotlyChart"], div[data-testid="stDataFrame"] {{
    background: #FFFFFF;
    border: 1px solid #E5EAE5;
    border-radius: 14px;
    padding: 10px 12px;
    box-shadow: 0 2px 10px rgba(27, 31, 42, 0.05);
}}

/* Buttons */
.stButton > button, .stDownloadButton > button {{
    background: {GREEN}; color: #FFFFFF; border: none; border-radius: 10px;
    font-weight: 600; padding: 0.55rem 1.1rem;
}}
.stButton > button:hover, .stDownloadButton > button:hover {{ background: {TEAL}; color: #FFFFFF; }}

/* Hero */
.hero {{
    background: linear-gradient(120deg, {GREEN} 0%, {TEAL} 60%, {LIME} 100%);
    border-radius: 20px; padding: 30px 34px; margin-bottom: 22px; color: #FFFFFF;
    box-shadow: 0 10px 30px rgba(30, 86, 49, 0.25);
}}
.hero .eyebrow {{ font-size: 0.78rem; font-weight: 700; letter-spacing: 0.14em; text-transform: uppercase; opacity: .85; }}
.hero h2 {{ font-size: 2.2rem; font-weight: 800; margin: 6px 0 8px 0; color: #FFFFFF; letter-spacing: -0.02em; }}
.hero p {{ font-size: 1.02rem; margin: 0 0 14px 0; opacity: .95; max-width: 820px; }}
.pill {{
    display: inline-block; background: rgba(255,255,255,0.18); border: 1px solid rgba(255,255,255,0.35);
    border-radius: 999px; padding: 4px 12px; margin: 0 6px 6px 0; font-size: 0.8rem; font-weight: 600;
}}

/* Sidebar brand */
.brand {{ display: flex; align-items: center; gap: 10px; margin: 4px 0 14px 0; }}
.brand .logo {{
    width: 40px; height: 40px; border-radius: 12px; display: flex; align-items: center; justify-content: center;
    background: linear-gradient(135deg, {GREEN}, {TEAL}); color: #fff; font-size: 1.3rem;
}}
.brand .name {{ font-weight: 800; color: {INK}; line-height: 1.1; }}
.brand .sub {{ font-size: 0.75rem; color: {GREY}; }}

/* Findings */
.finding {{
    display: flex; gap: 14px; align-items: flex-start; background: #FFFFFF;
    border: 1px solid #E5EAE5; border-radius: 12px; padding: 14px 16px; margin-bottom: 10px;
    box-shadow: 0 2px 8px rgba(27, 31, 42, 0.04);
}}
.finding .num {{
    min-width: 30px; height: 30px; border-radius: 50%; background: {GOLD}; color: #fff;
    font-weight: 800; display: flex; align-items: center; justify-content: center;
}}
.finding .txt {{ color: {INK}; font-size: 0.97rem; line-height: 1.5; }}

/* Footer */
.app-footer {{ text-align: center; color: {GREY}; font-size: 0.85rem; margin-top: 40px; padding-top: 18px; border-top: 1px solid #E5EAE5; }}
.app-footer a {{ color: {GREEN}; font-weight: 600; text-decoration: none; }}
</style>
""",
        unsafe_allow_html=True,
    )


def hero(title: str, subtitle: str, pills: list[str] | None = None, eyebrow: str = "Corporate carbon accounting"):
    pills_html = "".join(f'<span class="pill">{html.escape(p)}</span>' for p in (pills or []))
    st.markdown(
        f'<div class="hero"><div class="eyebrow">{html.escape(eyebrow)}</div>'
        f"<h2>{html.escape(title)}</h2><p>{html.escape(subtitle)}</p>{pills_html}</div>",
        unsafe_allow_html=True,
    )


def sidebar_brand(name: str = "GHG Inventory", sub: str = "Scope 1 · 2 · 3 dashboard", icon: str = "🌍"):
    st.sidebar.markdown(
        f'<div class="brand"><div class="logo">{icon}</div>'
        f'<div><div class="name">{html.escape(name)}</div><div class="sub">{html.escape(sub)}</div></div></div>',
        unsafe_allow_html=True,
    )


def findings_cards(items: list[str]):
    st.markdown("#### Key findings")
    st.markdown(
        "".join(
            f'<div class="finding"><div class="num">{i}</div><div class="txt">{html.escape(t)}</div></div>'
            for i, t in enumerate(items, 1)
        ),
        unsafe_allow_html=True,
    )


def footer():
    who = f'<a href="{AUTHOR_LINK}" target="_blank">{AUTHOR}</a>' if AUTHOR_LINK else f"<b>{AUTHOR}</b>"
    st.markdown(
        f'<div class="app-footer">Built by {who} · GHG Protocol Corporate Standard · '
        "Sample data is synthetic · Uploaded files are never stored</div>",
        unsafe_allow_html=True,
    )
