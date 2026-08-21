import io
import json
import os
import random
import re
import urllib.parse
import urllib.request
from dotenv import load_dotenv
import pandas as pd
from pydantic import BaseModel, Field
import streamlit as st

from google import genai
from google.genai import types

# Module openpyxl untuk Excel
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Module python-pptx untuk Slide Presentasi
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# Muat variabel dari file .env secara otomatis
load_dotenv()


# ==========================================
# 0. FUNGSI SANITIZER & DOWNLOAD GAMBAR
# ==========================================
def clean_text(text: str) -> str:
    """Mengekstrak karakter non-ASCII (seperti emoji) agar aman diproses."""
    if not text:
        return ""
    cleaned = re.sub(r"[^\x00-\x7F]+", "", text)
    return re.sub(r"\s+", " ", cleaned).strip()


def fetch_real_product_image(product_name: str) -> str | None:
    """Mencari foto produk asli via Wikipedia API."""
    try:
        clean_name = clean_text(product_name.split("(")[0])
        if not clean_name:
            return None

        search_url = f"https://en.wikipedia.org/w/api.php?action=query&generator=search&gsrsearch={urllib.parse.quote(clean_name)}&gsrlimit=1&prop=pageimages&pithumbsize=600&format=json"

        req = urllib.request.Request(
            search_url, headers={"User-Agent": "Mozilla/5.0"}
        )
        with urllib.request.urlopen(req, timeout=3) as response:
            data = json.loads(response.read().decode("utf-8"))
            pages = data.get("query", {}).get("pages", {})
            for page in pages.values():
                if "thumbnail" in page:
                    return page["thumbnail"]["source"]
    except Exception:
        return None
    return None


def download_image_bytes(url: str) -> io.BytesIO | None:
    """Mengunduh gambar dari URL untuk disisipkan ke slide PowerPoint."""
    if not url or not url.startswith("http"):
        return None
    try:
        req = urllib.request.Request(
            url, headers={"User-Agent": "Mozilla/5.0"}
        )
        with urllib.request.urlopen(req, timeout=4) as response:
            return io.BytesIO(response.read())
    except Exception:
        return None


# ==========================================
# 1. SCHEMA PYDANTIC
# ==========================================
class CompetitorItem(BaseModel):
    product_name: str = Field(description="Nama produk pembanding setara")
    origin_brand: str = Field(description="Asal negara atau brand produsen")
    origin_type: str = Field(
        description=(
            "Klasifikasi: 'Lokal / Indonesia' atau 'Luar Negeri / Impor'"
        )
    )
    alcohol_by_volume: str = Field(
        description="Kadar alkohol (% ABV, contoh: '16.9%', '40%')"
    )
    base_ingredients: str = Field(
        description="Bahan baku / komposisi utama produk pembanding"
    )
    price_point: str = Field(description="Estimasi rentang harga")
    key_difference: str = Field(
        description="Keunggulan atau perbedaan dibanding produk acuan"
    )
    product_weakness: str = Field(
        description="Kelemahan/kekurangan produk pembanding"
    )
    behind_story: str = Field(
        description=(
            "Cerita singkat di balik sejarah, filosofi, atau pembuatan produk"
            " pembanding"
        )
    )


class BusinessAnalysis(BaseModel):
    target_demographics: str = Field(description="Target konsumen")
    markup_margin_potential: str = Field(
        description="Potensi Margin/Markup di Bar/Resto"
    )
    suitable_venue: str = Field(description="Tipe tempat terbaik")
    investment_value: str = Field(description="Potensi investasi")


class BeverageItem(BaseModel):
    beverage_name: str = Field(description="Nama lengkap minuman")
    beverage_category: str = Field(description="Kategori spesifik")
    origin_type: str = Field(
        description=(
            "Klasifikasi: 'Lokal / Indonesia' atau 'Luar Negeri / Impor'"
        )
    )
    alcohol_by_volume: str = Field(
        description="Kadar alkohol (% ABV, contoh: '16.9%', '40%')"
    )
    unique_selling_point: str = Field(description="Unique Selling Point (USP)")
    behind_story: str = Field(
        description=(
            "Cerita singkat di balik asal-usul, sejarah, atau filosofi racikan"
            " produk (Short Behind Story)"
        )
    )
    product_weakness: str = Field(
        description="Kelemahan/kekurangan produk utama"
    )
    age_or_vintage: str = Field(description="Umur / Vintage")
    origin: str = Field(description="Negara & Region spesifik")
    base_ingredients: str = Field(description="Bahan baku / komposisi dasar")
    rating_scores: str = Field(description="Rating global")
    estimated_price: str = Field(description="Perkiraan harga retail")
    flavor_character: list[str] = Field(description="Karakter aroma & rasa")
    serving_recommendation: str = Field(description="Cara saji ideal")
    business_intelligence: BusinessAnalysis = Field(
        description="Analisis komersial"
    )
    equivalent_competitors: list[CompetitorItem] = Field(
        description="2-3 Produk pembanding/kompetitor yang setara di pasaran"
    )
    expert_note: str = Field(description="Catatan pakar")
    image_url: str = Field(default="")


class SelectorResponse(BaseModel):
    executive_summary: str = Field(description="Ringkasan eksekutif")
    selected_beverages: list[BeverageItem] = Field(
        description="Daftar minuman"
    )


SYSTEM_INSTRUCTION = """
Kamu adalah seorang Global Beverage Director & Master Sommelier.
Sortir minuman beralkohol terbaik berdasarkan kriteria input.
Wajib sertakan cerita singkat di balik latar belakang, sejarah, atau filosofi pembuatan produk (behind_story) untuk setiap produk utama dan pembandingnya.
Wajib analisis dan sebutkan KELEMAHAN / KEKURANGAN (product_weakness) secara objektif dari setiap produk utama dan pembandingnya.
Wajib sertakan persentase kadar alkohol (% ABV) yang akurat untuk setiap produk utama maupun produk pembanding setara.
Wajib berikan klasifikasi origin_type secara tegas apakah produk tersebut merupakan 'Lokal / Indonesia' atau 'Luar Negeri / Impor'.
Sebutkan bahan baku utama secara detail untuk setiap produk utama maupun produk pembanding setara.
Untuk setiap minuman, sertakan 2-3 produk pembanding/kompetitor yang setara di pasaran lengkap dengan behind story, kelemahan produk, kadar alkohol (% ABV), jenis kriteria asal, bahan baku, estimasi harga, dan perbedaan utamanya.
"""


# ==========================================
# 2. GENERATOR EXCEL DASHBOARD
# ==========================================
def create_excel_dashboard(data_json: dict) -> io.BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Executive Dashboard"
    ws.views.sheetView[0].showGridLines = True

    title_font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    title_fill = PatternFill(
        start_color="4C1D95", end_color="4C1D95", fill_type="solid"
    )

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(
        start_color="6D28D9", end_color="6D28D9", fill_type="solid"
    )

    sec_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    sec_fill = PatternFill(
        start_color="3B0764", end_color="3B0764", fill_type="solid"
    )

    thin_border = Border(
        left=Side(style="thin", color="E9D5FF"),
        right=Side(style="thin", color="E9D5FF"),
        top=Side(style="thin", color="E9D5FF"),
        bottom=Side(style="thin", color="E9D5FF"),
    )

    ws.merge_cells("A1:T2")
    banner = ws["A1"]
    banner.value = "GLOBAL BEVERAGE & BENCHMARKING INTELLIGENCE DASHBOARD"
    banner.font = title_font
    banner.fill = title_fill
    banner.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A4:T6")
    summary_cell = ws["A4"]
    exec_text = data_json.get("executive_summary", "")
    summary_cell.value = f"📌 EXECUTIVE STRATEGY SUMMARY:\n{exec_text}"
    summary_cell.font = Font(name="Calibri", size=10, italic=True)
    summary_cell.fill = PatternFill(
        start_color="F3E8FF", end_color="F3E8FF", fill_type="solid"
    )
    summary_cell.alignment = Alignment(
        vertical="center", horizontal="left", wrap_text=True
    )

    ws_chart_data = wb.create_sheet(title="ChartData")
    ws_chart_data["A1"] = "Kategori"
    ws_chart_data["B1"] = "Jumlah Produk"

    items = data_json.get("selected_beverages", [])
    cat_counts = {}
    for item in items:
        cat = item.get("beverage_category", "Lainnya")
        cat_counts[cat] = cat_counts.get(cat, 0) + 1

    r = 2
    for cat, count in cat_counts.items():
        ws_chart_data.cell(row=r, column=1, value=cat)
        ws_chart_data.cell(row=r, column=2, value=count)
        r += 1

    labels_ref = Reference(ws_chart_data, min_col=1, min_row=2, max_row=r - 1)
    data_ref = Reference(ws_chart_data, min_col=2, min_row=1, max_row=r - 1)

    pie = PieChart()
    pie.add_data(data_ref, titles_from_data=True)
    pie.set_categories(labels_ref)
    pie.title = "Proporsi Kategori Minuman"
    pie.width = 14
    pie.height = 7.5
    ws.add_chart(pie, "A8")

    bar = BarChart()
    bar.type = "col"
    bar.style = 10
    bar.title = "Sebaran Produk per Kategori"
    bar.y_axis.title = "Jumlah Botol"
    bar.add_data(data_ref, titles_from_data=True)
    bar.set_categories(labels_ref)
    bar.width = 16
    bar.height = 7.5
    ws.add_chart(bar, "I8")

    start_row = 23
    ws.merge_cells(f"A{start_row-1}:T{start_row-1}")
    sec_title = ws[f"A{start_row-1}"]
    sec_title.value = "📊 DETAILED BEVERAGE ANALYSIS & COMPETITOR BENCHMARKING"
    sec_title.font = sec_font
    sec_title.fill = sec_fill
    sec_title.alignment = Alignment(horizontal="left", vertical="center")

    headers = [
        "No",
        "Nama Minuman",
        "Kategori",
        "Klasifikasi Asal",
        "Kadar Alkohol (ABV)",
        "Umur / Vintage",
        "Asal (Origin)",
        "Bahan Baku Utama",
        "Behind Story (Sejarah)",
        "Rating Global",
        "Estimasi Harga",
        "Unique Selling Point (USP)",
        "Kelemahan Produk",
        "Tasting Notes",
        "Cara Saji",
        "Produk Pembanding Setara (ABV, Bahan & Story)",
        "Margin Profit",
        "Target Market",
        "Venue Cocok",
        "Potensi Investasi",
    ]

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=start_row, column=col_idx)
        cell.value = h
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = thin_border

    for idx, item in enumerate(items, 1):
        curr_row = start_row + idx
        biz = item.get("business_intelligence", {})
        comps = item.get("equivalent_competitors", [])

        comp_str = "\n".join(
            [
                (
                    f"• {c.get('product_name')} ({c.get('origin_type')} -"
                    f" {c.get('origin_brand')}) | ABV:"
                    f" {c.get('alcohol_by_volume')} | Story:"
                    f" {c.get('behind_story')} | Kelemahan:"
                    f" {c.get('product_weakness')} | Bahan:"
                    f" {c.get('base_ingredients')} | Harga:"
                    f" {c.get('price_point')}"
                )
                for c in comps
            ]
        )

        row_values = [
            idx,
            item.get("beverage_name"),
            item.get("beverage_category"),
            item.get("origin_type"),
            item.get("alcohol_by_volume"),
            item.get("age_or_vintage"),
            item.get("origin"),
            item.get("base_ingredients"),
            item.get("behind_story"),
            item.get("rating_scores"),
            item.get("estimated_price"),
            item.get("unique_selling_point"),
            item.get("product_weakness"),
            ", ".join(item.get("flavor_character", [])),
            item.get("serving_recommendation"),
            comp_str,
            biz.get("markup_margin_potential"),
            biz.get("target_demographics"),
            biz.get("suitable_venue"),
            biz.get("investment_value"),
        ]

        for col_idx, val in enumerate(row_values, 1):
            cell = ws.cell(row=curr_row, column=col_idx)
            cell.value = val
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

            if idx % 2 == 0:
                cell.fill = PatternFill(
                    start_color="FAF5FF", end_color="FAF5FF", fill_type="solid"
                )

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.row >= start_row:
                val_str = str(cell.value or "")
                if len(val_str) > max_len:
                    max_len = len(val_str)
        ws.column_dimensions[col_letter].width = min(
            max(max_len + 3, 12), 35
        )

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ==========================================
# 3. DYNAMIC POWERPOINT GENERATOR (HIGH-END DESIGN)
# ==========================================
def create_pptx_presentation(data_json: dict) -> io.BytesIO:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    THEME_PALETTES = [
        {
            "name": "Luxury Dark Gold",
            "card_bg": RGBColor(15, 23, 42),
            "border": RGBColor(234, 179, 8),
            "primary": RGBColor(250, 204, 21),
            "accent": RGBColor(168, 85, 247),
            "text": RGBColor(241, 245, 249),
            "subtext": RGBColor(203, 213, 225),
            "card_header": RGBColor(30, 41, 59),
        },
        {
            "name": "Royal Emerald Sommelier",
            "card_bg": RGBColor(6, 44, 34),
            "border": RGBColor(16, 185, 129),
            "primary": RGBColor(110, 231, 183),
            "accent": RGBColor(253, 224, 71),
            "text": RGBColor(240, 253, 244),
            "subtext": RGBColor(187, 247, 208),
            "card_header": RGBColor(4, 120, 87),
        },
        {
            "name": "Midnight Wine Crimson",
            "card_bg": RGBColor(45, 10, 24),
            "border": RGBColor(244, 63, 94),
            "primary": RGBColor(251, 146, 60),
            "accent": RGBColor(253, 186, 116),
            "text": RGBColor(255, 241, 242),
            "subtext": RGBColor(254, 205, 211),
            "card_header": RGBColor(136, 19, 55),
        },
        {
            "name": "Cyber Obsidian Cyan",
            "card_bg": RGBColor(18, 24, 38),
            "border": RGBColor(14, 165, 233),
            "primary": RGBColor(56, 189, 248),
            "accent": RGBColor(129, 140, 248),
            "text": RGBColor(240, 249, 255),
            "subtext": RGBColor(186, 230, 253),
            "card_header": RGBColor(3, 105, 161),
        },
    ]

    COVER_BG_POOL = [
        (
            "https://images.unsplash.com/photo-1514933651103-005eec06c04b?w=1600&q=80"
        ),
        (
            "https://images.unsplash.com/photo-1510812431401-41d2bd2722f3?w=1600&q=80"
        ),
        (
            "https://images.unsplash.com/photo-1551024709-8f23befc6f87?w=1600&q=80"
        ),
    ]
    PRODUCT_BG_POOL = [
        (
            "https://images.unsplash.com/photo-1527281400683-1aae777175f8?w=1600&q=80"
        ),
        (
            "https://images.unsplash.com/photo-1563227812-0ea4c22e6cc8?w=1600&q=80"
        ),
        (
            "https://images.unsplash.com/photo-1574096079513-d8259312b785?w=1600&q=80"
        ),
        (
            "https://images.unsplash.com/photo-1517457373958-b7bdd4587205?w=1600&q=80"
        ),
    ]

    theme = random.choice(THEME_PALETTES)
    cover_bg_url = random.choice(COVER_BG_POOL)

    def set_slide_bg(slide, image_url: str):
        img_bytes = download_image_bytes(image_url)
        if img_bytes:
            slide.shapes.add_picture(
                img_bytes, 0, 0, Inches(13.333), Inches(7.5)
            )
        else:
            bg = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5)
            )
            bg.fill.solid()
            bg.fill.fore_color.rgb = theme["card_bg"]
            bg.line.fill.background()

    def add_modern_header(slide, title_text: str):
        banner = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.1)
        )
        banner.fill.solid()
        banner.fill.fore_color.rgb = theme["card_header"]
        banner.line.color.rgb = theme["border"]

        accent_line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, 0, Inches(1.05), Inches(13.333), Inches(0.05)
        )
        accent_line.fill.solid()
        accent_line.fill.fore_color.rgb = theme["border"]
        accent_line.line.fill.background()

        tb = slide.shapes.add_textbox(
            Inches(0.8), Inches(0.18), Inches(11.5), Inches(0.7)
        )
        p = tb.text_frame.paragraphs[0]
        p.text = title_text
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = theme["primary"]

    # SLIDE 1: COVER
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide1, cover_bg_url)

    overlay1 = slide1.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1.0),
        Inches(1.5),
        Inches(11.333),
        Inches(4.5),
    )
    overlay1.fill.solid()
    overlay1.fill.fore_color.rgb = theme["card_bg"]
    overlay1.line.color.rgb = theme["border"]
    overlay1.line.width = Pt(2)

    tf_c = overlay1.text_frame
    tf_c.word_wrap = True

    p1 = tf_c.paragraphs[0]
    p1.text = "GLOBAL BEVERAGE INTELLIGENCE"
    p1.font.size = Pt(36)
    p1.font.bold = True
    p1.font.color.rgb = theme["primary"]
    p1.alignment = PP_ALIGN.CENTER

    p2 = tf_c.add_paragraph()
    p2.text = "Executive Strategy & Competitor Benchmarking Deck\n"
    p2.font.size = Pt(20)
    p2.font.color.rgb = theme["text"]
    p2.alignment = PP_ALIGN.CENTER

    p3 = tf_c.add_paragraph()
    p3.text = (
        f"• Theme Edition: {theme['name']}  • AI Analytics  • HORECA Commercial"
        " Strategy"
    )
    p3.font.size = Pt(13)
    p3.font.color.rgb = theme["accent"]
    p3.alignment = PP_ALIGN.CENTER

    # SLIDE 2: EXECUTIVE SUMMARY
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_bg(slide2, cover_bg_url)
    add_modern_header(slide2, "📋 Executive Strategy Summary")

    card_sum = slide2.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.8),
        Inches(1.5),
        Inches(11.733),
        Inches(5.3),
    )
    card_sum.fill.solid()
    card_sum.fill.fore_color.rgb = theme["card_bg"]
    card_sum.line.color.rgb = theme["border"]

    tf_s = card_sum.text_frame
    tf_s.word_wrap = True

    ps1 = tf_s.paragraphs[0]
    ps1.text = "STRATEGI PENYORTIRAN & NILAI KOMERSIAL HORECA\n"
    ps1.font.size = Pt(18)
    ps1.font.bold = True
    ps1.font.color.rgb = theme["primary"]

    ps2 = tf_s.add_paragraph()
    ps2.text = data_json.get("executive_summary", "")
    ps2.font.size = Pt(14)
    ps2.font.color.rgb = theme["text"]

    # SLIDES PER PRODUK
    beverages = data_json.get("selected_beverages", [])
    for idx, item in enumerate(beverages, 1):
        product_bg_url = random.choice(PRODUCT_BG_POOL)

        # SUB-SLIDE A
        slide_a = prs.slides.add_slide(blank_layout)
        set_slide_bg(slide_a, product_bg_url)
        add_modern_header(
            slide_a,
            f"#{idx} {item.get('beverage_name')} [{item.get('origin_type')}] —"
            " 📌 SPESIFIKASI UTAMA",
        )

        img_card = slide_a.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.8),
            Inches(1.5),
            Inches(5.2),
            Inches(5.3),
        )
        img_card.fill.solid()
        img_card.fill.fore_color.rgb = theme["card_bg"]
        img_card.line.color.rgb = theme["border"]

        img_url = item.get("image_url") or fetch_real_product_image(
            item.get("beverage_name", "")
        )
        img_bytes = download_image_bytes(img_url) if img_url else None

        if img_bytes:
            try:
                slide_a.shapes.add_picture(
                    img_bytes,
                    Inches(1.1),
                    Inches(1.8),
                    width=Inches(4.6),
                    height=Inches(4.7),
                )
            except Exception:
                pass

        info_card = slide_a.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(6.3),
            Inches(1.5),
            Inches(6.2),
            Inches(5.3),
        )
        info_card.fill.solid()
        info_card.fill.fore_color.rgb = theme["card_bg"]
        info_card.line.color.rgb = theme["border"]

        tf_ia = info_card.text_frame
        tf_ia.word_wrap = True

        p_a1 = tf_ia.paragraphs[0]
        p_a1.text = "DETAIL SPESIFIKASI PRODUK\n"
        p_a1.font.bold = True
        p_a1.font.size = Pt(18)
        p_a1.font.color.rgb = theme["primary"]

        spec_text = (
            f"• Nama Lengkap:  {item.get('beverage_name')}\n"
            f"• Klasifikasi Asal:  {item.get('origin_type')}\n"
            f"• Kadar Alkohol (ABV):  {item.get('alcohol_by_volume')}\n"
            f"• Kategori Spesifik:  {item.get('beverage_category')}\n"
            f"• Wilayah Asal:  {item.get('origin')}\n"
            f"• Bahan Baku Utama:  {item.get('base_ingredients')}\n"
            f"• Rating Global:  {item.get('rating_scores')}\n"
            f"• Perkiraan Harga:  {item.get('estimated_price')}"
        )
        p_a2 = tf_ia.add_paragraph()
        p_a2.text = spec_text
        p_a2.font.size = Pt(12)
        p_a2.font.color.rgb = theme["text"]

        # SUB-SLIDE B
        slide_b = prs.slides.add_slide(blank_layout)
        set_slide_bg(slide_b, product_bg_url)
        add_modern_header(
            slide_b,
            f"#{idx} {item.get('beverage_name')} — ⚔️ BENCHMARKING & BEHIND"
            " STORY",
        )

        card_b1 = slide_b.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.8),
            Inches(1.5),
            Inches(5.7),
            Inches(5.3),
        )
        card_b1.fill.solid()
        card_b1.fill.fore_color.rgb = theme["card_bg"]
        card_b1.line.color.rgb = theme["border"]

        tf_b1 = card_b1.text_frame
        tf_b1.word_wrap = True

        pb1_h = tf_b1.paragraphs[0]
        pb1_h.text = "📖 BEHIND STORY & USP\n"
        pb1_h.font.bold = True
        pb1_h.font.size = Pt(16)
        pb1_h.font.color.rgb = theme["primary"]

        tasting_str = ", ".join(item.get("flavor_character", []))
        b1_text = (
            f"📖 Behind Story:\n{item.get('behind_story')}\n\n"
            f"🔥 Unique Selling Point (USP):\n{item.get('unique_selling_point')}\n\n"
            f"⚠️ Kelemahan Utama:\n{item.get('product_weakness')}\n\n"
            f"👃 Tasting Notes: {tasting_str}"
        )
        pb1_body = tf_b1.add_paragraph()
        pb1_body.text = b1_text
        pb1_body.font.size = Pt(11)
        pb1_body.font.color.rgb = theme["text"]

        card_b2 = slide_b.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(6.8),
            Inches(1.5),
            Inches(5.7),
            Inches(5.3),
        )
        card_b2.fill.solid()
        card_b2.fill.fore_color.rgb = theme["card_bg"]
        card_b2.line.color.rgb = theme["border"]

        tf_b2 = card_b2.text_frame
        tf_b2.word_wrap = True

        pb2_h = tf_b2.paragraphs[0]
        pb2_h.text = "⚔️ PRODUK PEMBANDING SETARA\n"
        pb2_h.font.bold = True
        pb2_h.font.size = Pt(16)
        pb2_h.font.color.rgb = theme["primary"]

        comps = item.get("equivalent_competitors", [])
        comp_body_text = ""
        for c in comps:
            comp_body_text += (
                f"• {c.get('product_name')} [{c.get('origin_type')}]\n"
                f"  Behind Story: {c.get('behind_story')}\n"
                f"  ABV: {c.get('alcohol_by_volume')} | Harga:"
                f" {c.get('price_point')}\n"
                f"  Bedanya: {c.get('key_difference')}\n"
                f"  Kelemahan: {c.get('product_weakness')}\n\n"
            )

        pb2_body = tf_b2.add_paragraph()
        pb2_body.text = (
            comp_body_text if comp_body_text else "Tidak ada data pembanding."
        )
        pb2_body.font.size = Pt(10)
        pb2_body.font.color.rgb = theme["subtext"]

        # SUB-SLIDE C
        slide_c = prs.slides.add_slide(blank_layout)
        set_slide_bg(slide_c, product_bg_url)
        add_modern_header(
            slide_c,
            f"#{idx} {item.get('beverage_name')} — 💼 ANALISIS BISNIS &"
            " KOMERSIAL HORECA",
        )

        biz = item.get("business_intelligence", {})
        biz_boxes = [
            (
                "📈 Potensi Margin Profit",
                biz.get("markup_margin_potential"),
                Inches(0.8),
                Inches(1.5),
            ),
            (
                "👥 Target Konsumen",
                biz.get("target_demographics"),
                Inches(6.8),
                Inches(1.5),
            ),
            (
                "🏢 Venue Terbaik",
                biz.get("suitable_venue"),
                Inches(0.8),
                Inches(4.3),
            ),
            (
                "🔄 Potensi Investasi / Stok",
                biz.get("investment_value"),
                Inches(6.8),
                Inches(4.3),
            ),
        ]

        for b_title, b_val, left_x, top_y in biz_boxes:
            b_shape = slide_c.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                left_x,
                top_y,
                Inches(5.7),
                Inches(2.5),
            )
            b_shape.fill.solid()
            b_shape.fill.fore_color.rgb = theme["card_bg"]
            b_shape.line.color.rgb = theme["border"]

            tf_bx = b_shape.text_frame
            tf_bx.word_wrap = True

            p_bh = tf_bx.paragraphs[0]
            p_bh.text = b_title
            p_bh.font.bold = True
            p_bh.font.size = Pt(16)
            p_bh.font.color.rgb = theme["primary"]

            p_bv = tf_bx.add_paragraph()
            p_bv.text = b_val or "-"
            p_bv.font.size = Pt(13)
            p_bv.font.color.rgb = theme["text"]

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer


# ==========================================
# 4. STREAMLIT APP CONFIG & FORM
# ==========================================
st.set_page_config(
    page_title="Global Beverage & Business AI Pro", page_icon="🥃", layout="wide"
)

st.title("🍾 Global Beverage Business")
st.write(
    "Beverages Analysis, Sorting & Equivalent Product Benchmarking System"
    " Based on AI."
)

# DETEKSI AUTOMATIS KUNCI API
# Mengambil dari Streamlit Secrets atau variabel lingkungan .env
auto_api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get(
    "GEMINI_API_KEY", ""
)

if auto_api_key:
    st.sidebar.success("✅ API Key Terdeteksi Otomatis!")
    raw_api_key = st.sidebar.text_input(
        "Gemini API Key (Opsional / Override):",
        type="password",
        value=auto_api_key,
    )
else:
    st.sidebar.warning("⚠️ API Key tidak terdeteksi di .env / Secrets")
    raw_api_key = st.sidebar.text_input(
        "Masukkan Gemini API Key:",
        type="password",
        value="",
    )

api_key = raw_api_key.strip()

with st.form("form_analisis"):
    col1, col2 = st.columns(2)
    with col1:
        beverage_type = st.selectbox(
            "🥃 Beverages Category:",
            [
                "🌐 Universal / Bebas (Campuran Terbaik)",
                "🍷 Wine",
                "🥃 Whisky / Whiskey",
                "🍺 Beer / Craft Beer",
                "🍶 Sake",
                "🇰🇷 Soju",
                "🍸 Gin",
                "🧊 Vodka",
                "🌵 Tequila / Mezcal",
                "🏴‍☠️ Rum",
                "☠ Brandy",
            ],
        )
        origin_scope = st.selectbox(
            "🇲🇨 Scope Of Origin:",
            [
                "🌐 Bebas / Campuran (Lokal & Impor)",
                "🇮🇩 Khusus Lokal / Indonesia",
                "🌍 Khusus Impor / Luar Negeri",
                "⚔️ Komparasi Head-to-Head (Lokal vs Impor)",
            ],
        )
        brand_input = st.text_input(
            "🏷️ Brand / Product (Optional):",
            placeholder="Contoh: Kura Kura Beer / Jinro Fresh",
        )
        item_count = st.selectbox(
            "🔢 Sort By Range:",
            options=list(range(1, 16)),
            index=2,
        )
    with col2:
        sort_priority = st.selectbox(
            "🎯 Main Priority:",
            [
                "🌐 Universal / Seimbang",
                "⭐ Rating Global Tertinggi",
                "💰 Margin Profit HORECA Tertinggi",
                "⚖️ Best Value",
            ],
        )
        venue_type = st.selectbox(
            "🏢 Type Of Venue:",
            [
                "🌐 Universal / Semua Tipe",
                "🍸 Speakeasy Bar",
                "🍷 Fine Dining",
                "🍻 Casual Pub",
                "🪩 Nightclub",
            ],
        )
        budget_category = st.selectbox(
            "💵 Price Category:",
            [
                "🌐 Universal / Semua Range",
                "🟢 Entry Level (< $30)",
                "🟡 Premium Pour ($30 - $100)",
                "🟠 Top Shelf ($100 - $300)",
                "🔴 Collector (> $300)",
            ],
        )
        origin_input = st.text_input("🌍 Specific Of Origin (Optional):")

    btn_process = st.form_submit_button("🔍 Punch Down", type="primary")

# ==========================================
# 5. EXECUTION & DISPLAY
# ==========================================
if btn_process:
    if not api_key:
        st.error(
            "❌ API Key tidak terdeteksi. Silakan atur file .env atau masukkan"
            " manual di sidebar!"
        )
    else:
        c_bev = clean_text(beverage_type)
        c_scope = clean_text(origin_scope)
        c_prio = clean_text(sort_priority)
        c_ven = clean_text(venue_type)
        c_bud = clean_text(budget_category)

        prompt_query = (
            f"Sortir {item_count} minuman: Kategori: {c_bev}, Scope Asal:"
            f" {c_scope}, Brand/Acuan: {brand_input or 'Bebas'}, Prioritas:"
            f" {c_prio}, Venue: {c_ven}, Budget: {c_bud}, Region:"
            f" {origin_input or 'Bebas'}. Tentukan cerita singkat latar belakang"
            " pembuatan (behind_story), kelemahan produk (product_weakness),"
            " kriteria origin_type ('Lokal / Indonesia' atau 'Luar Negeri /"
            " Impor') serta kadar alkohol (% ABV) secara presisi."
        )

        with st.spinner(
            "📊 Menganalisis & mencari pembanding setara dari database"
            " global..."
        ):
            try:
                client = genai.Client(api_key=api_key)
                try:
                    response = client.models.generate_content(
                        model="gemini-3.6-flash",
                        contents=prompt_query,
                        config=types.GenerateContentConfig(
                            system_instruction=SYSTEM_INSTRUCTION,
                            temperature=0.2,
                            response_mime_type="application/json",
                            response_schema=SelectorResponse,
                        ),
                    )
                except Exception:
                    response = client.models.generate_content(
                        model="gemini-3.5-flash-lite",
                        contents=prompt_query,
                        config=types.GenerateContentConfig(
                            system_instruction=SYSTEM_INSTRUCTION,
                            temperature=0.2,
                            response_mime_type="application/json",
                            response_schema=SelectorResponse,
                        ),
                    )

                data = json.loads(response.text)

                st.success(
                    "✅ Analisis & Benchmarking"
                    f" {len(data.get('selected_beverages', []))} Minuman"
                    " Selesai!"
                )

                excel_buffer = create_excel_dashboard(data)
                pptx_buffer = create_pptx_presentation(data)

                st.subheader("📥 Export Hasil Laporan & Benchmarking")
                col_btn1, col_btn2 = st.columns(2)

                with col_btn1:
                    st.download_button(
                        label="📊 Download Excel Dashboard (.xlsx)",
                        data=excel_buffer,
                        file_name="Executive_Beverage_Dashboard.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        ),
                        type="primary",
                        use_container_width=True,
                    )

                with col_btn2:
                    st.download_button(
                        label="🖥️ Download PowerPoint Presentation (.pptx)",
                        data=pptx_buffer,
                        file_name="Executive_Beverage_Presentation.pptx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        ),
                        type="primary",
                        use_container_width=True,
                    )

                st.markdown("---")
                st.subheader("📋 Ringkasan Laporan Executive")
                st.info(data["executive_summary"])
                st.markdown("---")

                # Tampilan UI Streamlit
                for i, bev in enumerate(data["selected_beverages"], 1):
                    biz = bev["business_intelligence"]

                    origin_tag = (
                        "🇮🇩 LOKAL"
                        if "lokal" in bev.get("origin_type", "").lower()
                        or "indonesia" in bev.get("origin_type", "").lower()
                        else "🌍 IMPOR"
                    )

                    st.markdown(
                        f"### #{i} {bev['beverage_name']}"
                        f" ({bev['age_or_vintage']}) &nbsp; `{origin_tag}`"
                        f" &nbsp; 🧪 `{bev['alcohol_by_volume']}`"
                    )

                    c1, c2, c3 = st.columns([1.5, 2, 2])
                    with c1:
                        st.markdown(
                            f"**🥃 Kategori:** {bev['beverage_category']}\n\n"
                            f"**🏷️ Klasifikasi:** `{bev['origin_type']}`\n\n"
                            f"**🧪 Kadar Alkohol:** `{bev['alcohol_by_volume']}`\n\n"
                            f"**🌾 Bahan Baku:** {bev['base_ingredients']}\n\n"
                            f"**📍 Asal:** {bev['origin']}\n\n"
                            f"**⭐ Rating:** `{bev['rating_scores']}`\n\n"
                            f"**💵 Harga:** `{bev['estimated_price']}`"
                        )
                    with c2:
                        st.markdown(
                            f"**📖 Behind Story:** {bev['behind_story']}\n\n"
                            f"**🔥 USP:** {bev['unique_selling_point']}\n\n"
                            f"**⚠️ Kelemahan:** {bev['product_weakness']}\n\n"
                            f"**👃 Notes:** {', '.join(bev['flavor_character'])}\n\n"
                            f"**🧊 Saji:** {bev['serving_recommendation']}"
                        )
                    with c3:
                        st.markdown(
                            f"**📈 Margin:** {biz['markup_margin_potential']}\n\n"
                            f"**👥 Target:** {biz['target_demographics']}\n\n"
                            f"**🏢 Venue:** {biz['suitable_venue']}"
                        )

                    st.markdown(
                        "#### ⚔️ Produk Pembanding & Benchmarking Setara"
                    )
                    comps = bev.get("equivalent_competitors", [])
                    if comps:
                        comp_cols = st.columns(len(comps))
                        for idx_c, comp in enumerate(comps):
                            comp_tag = (
                                "🇮🇩 LOKAL"
                                if "lokal"
                                in comp.get("origin_type", "").lower()
                                or "indonesia"
                                in comp.get("origin_type", "").lower()
                                else "🌍 IMPOR"
                            )
                            with comp_cols[idx_c]:
                                st.info(
                                    f"**{comp['product_name']}** `{comp_tag}`\n\n"
                                    "📖 **Behind Story:**"
                                    f" {comp['behind_story']}\n\n"
                                    f"🧪 **ABV:** `{comp['alcohol_by_volume']}`\n\n"
                                    f"📍 **Asal/Brand:** {comp['origin_brand']}\n\n"
                                    "🌾 **Bahan Baku:**"
                                    f" {comp['base_ingredients']}\n\n"
                                    f"💵 **Harga:** `{comp['price_point']}`\n\n"
                                    f"💡 **Bedanya:** {comp['key_difference']}\n\n"
                                    f"⚠️ **Kelemahan:** {comp['product_weakness']}"
                                )
                    st.markdown("---")

            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")
