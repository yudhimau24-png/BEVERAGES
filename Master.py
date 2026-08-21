import io
import json
import os
import random
import re
import urllib.parse
import urllib.request
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

# Import dotenv secara aman
try:
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass


# ==============================================================================
# 1. GROUND TRUTH MASTER REGISTRY (DATABASE RIIL INDONESIA & GLOBAL)
# ==============================================================================
MASTER_VERIFIED_BRANDS = {
    # --- PT BALARAJA BARAT INDAH ---
    "kawa": {
        "real_name": "Anggur Kawa-Kawa",
        "producer": "PT Balaraja Barat Indah (PT BBI)",
        "origin_type": "Lokal / Indonesia",
        "abv": "19.8%",
        "behind_story": "Diproduksi secara mandiri oleh PT Balaraja Barat Indah di Balaraja, Tangerang. Merupakan lini anggur lokal manis beralkohol tinggi untuk segmen komersial.",
    },
    "alexis": {
        "real_name": "Anggur Alexis",
        "producer": "PT Balaraja Barat Indah (PT BBI)",
        "origin_type": "Lokal / Indonesia",
        "abv": "19.8%",
        "behind_story": "Diproduksi oleh PT Balaraja Barat Indah (produsen yang sama dengan Kawa-Kawa). Alexis hadir sebagai lini anggur merah & putih modern dengan cita rasa manis pekat.",
    },
    "wija": {
        "real_name": "Wija Soju",
        "producer": "PT Balaraja Barat Indah (PT BBI)",
        "origin_type": "Lokal / Indonesia",
        "abv": "15.0%",
        "behind_story": "Lini produk Soju lokal dari PT Balaraja Barat Indah dengan varian rasa buah yang disesuaikan dengan selera pasar anak muda Indonesia.",
    },
    # --- PT ASTIDAMA ADIMUKTI ---
    "mix max": {
        "real_name": "Mix Max Exotic Flavors",
        "producer": "PT Astidama Adimukti",
        "origin_type": "Lokal / Indonesia",
        "abv": "4.8%",
        "behind_story": "Diproduksi oleh PT Astidama Adimukti. Merupakan pionir Alco Pop / RTD (Ready To Drink) berbasis vodka rasa buah paling populer di Indonesia.",
    },
    "iceland": {
        "real_name": "Iceland Vodka",
        "producer": "PT Astidama Adimukti",
        "origin_type": "Lokal / Indonesia",
        "abv": "40.0%",
        "behind_story": "Diproduksi oleh PT Astidama Adimukti di Indonesia berbasis olahan gandum terdistilasi presisi untuk pasar vodka jernih entry-level.",
    },
    "drum": {
        "real_name": "Drum Whisky",
        "producer": "PT Astidama Adimukti",
        "origin_type": "Lokal / Indonesia",
        "abv": "43.0%",
        "behind_story": "Pelopor Whisky lokal buatan PT Astidama Adimukti yang dimatangkan dengan potongan kayu ek untuk menghasilkan aroma kayu alami.",
    },
    "mansion house": {
        "real_name": "Mansion House Whisky/Brandy",
        "producer": "PT Astidama Adimukti",
        "origin_type": "Lokal / Indonesia",
        "abv": "40.0%",
        "behind_story": "Merek spirit lokal legendaris di bawah naungan PT Astidama Adimukti untuk segmen pasar komersial luas.",
    },
    "empire": {
        "real_name": "Empire Gin",
        "producer": "PT Astidama Adimukti",
        "origin_type": "Lokal / Indonesia",
        "abv": "40.0%",
        "behind_story": "Gin lokal terjangkau buatan PT Astidama Adimukti yang mendominasi pasar racikan bar.",
    },
    # --- OT GROUP ---
    "amer ot": {
        "real_name": "Anggur Merah Cap Orang Tua (Amer OT)",
        "producer": "OT Group (Orang Tua Group)",
        "origin_type": "Lokal / Indonesia",
        "abv": "14.7%",
        "behind_story": "Diproduksi oleh OT Group sejak tahun 1948. Pelopor anggur ramuan tradisional Indonesia dari fermentasi anggur dan herbal lokal.",
    },
    "intisari": {
        "real_name": "Intisari Ginseng",
        "producer": "OT Group (Orang Tua Group)",
        "origin_type": "Lokal / Indonesia",
        "abv": "19.8%",
        "behind_story": "Minuman herbal obat tradisional beralkohol berbasis ekstrak ginseng buatan OT Group.",
    },
    # --- PRODUSEN BIR & WINE INDONESIA ---
    "bintang": {
        "real_name": "Bir Bintang",
        "producer": "PT Multi Bintang Indonesia Tbk",
        "origin_type": "Lokal / Indonesia",
        "abv": "4.7%",
        "behind_story": "Diproduksi oleh PT Multi Bintang Indonesia Tbk (Heineken Group) sejak 1931. Merek pilsner paling populer di Indonesia.",
    },
    "anker": {
        "real_name": "Anker Beer",
        "producer": "PT Delta Djakarta Tbk",
        "origin_type": "Lokal / Indonesia",
        "abv": "4.5%",
        "behind_story": "Diproduksi oleh PT Delta Djakarta Tbk sejak 1932 sebagai salah satu bir pilsner tertua di tanah air.",
    },
    "sababay": {
        "real_name": "Sababay Wine",
        "producer": "PT Sababay Industry",
        "origin_type": "Lokal / Indonesia",
        "abv": "13.0%",
        "behind_story": "Produsen winery lokal berbasis di Buleleng, Bali yang bermitra dengan petani anggur lokal Bali.",
    },
}

# --- DATABASE CADANGAN: PENYELAMAT JIKA AI KURANG MEMBERI KOMPETITOR ---
REAL_COMPETITOR_FALLBACKS = {
    "alco pops": [
        {
            "product_name": "Mix Max Exotic Flavors",
            "origin_brand": "PT Astidama Adimukti (Lokal / Indonesia)",
            "origin_type": "Lokal / Indonesia",
            "alcohol_by_volume": "4.8%",
            "base_ingredients": "Vodka netral, perisa buah, air berkarbonasi",
            "price_point": "IDR 25,000 - IDR 35,000",
            "key_difference": "Alco pop lokal dengan variasi rasa buah manis",
            "product_weakness": "Kandungan gula cukup tinggi",
            "behind_story": "Pionir RTD lokal buatan PT Astidama Adimukti."
        },
        {
            "product_name": "Smirnoff Ice",
            "origin_brand": "Diageo (Luar Negeri / Impor)",
            "origin_type": "Luar Negeri / Impor",
            "alcohol_by_volume": "4.5%",
            "base_ingredients": "Smirnoff Vodka, perisa lemon, soda",
            "price_point": "IDR 35,000 - IDR 50,000",
            "key_difference": "Rasa lemon yang ikonik dan brand global",
            "product_weakness": "Harga lebih mahal dibanding RTD lokal",
            "behind_story": "Merek RTD global paling laris."
        }
    ],
    "whisky": [
        {
            "product_name": "Drum Whisky",
            "origin_brand": "PT Astidama Adimukti (Lokal / Indonesia)",
            "origin_type": "Lokal / Indonesia",
            "alcohol_by_volume": "43.0%",
            "base_ingredients": "Malt, spirit lokal, ekstrak oak",
            "price_point": "IDR 150,000 - IDR 200,000",
            "key_difference": "Whisky lokal harga sangat terjangkau",
            "product_weakness": "Aroma alkohol cukup tajam",
            "behind_story": "Diproduksi oleh PT Astidama Adimukti sebagai andalan whisky lokal."
        },
        {
            "product_name": "Jim Beam White Label",
            "origin_brand": "Beam Suntory (Luar Negeri / Impor)",
            "origin_type": "Luar Negeri / Impor",
            "alcohol_by_volume": "40.0%",
            "base_ingredients": "Jagung, rye, barley",
            "price_point": "IDR 450,000 - IDR 600,000",
            "key_difference": "Bourbon asli yang dimatangkan di tong oak baru",
            "product_weakness": "Harga terpengaruh oleh pajak impor",
            "behind_story": "Kentucky Straight Bourbon ikonik global."
        }
    ],
    "vodka": [
        {
            "product_name": "Iceland Vodka",
            "origin_brand": "PT Astidama Adimukti (Lokal / Indonesia)",
            "origin_type": "Lokal / Indonesia",
            "alcohol_by_volume": "40.0%",
            "base_ingredients": "Gandum lokal terdistilasi",
            "price_point": "IDR 130,000 - IDR 170,000",
            "key_difference": "Vodka lokal dengan harga paling ekonomis",
            "product_weakness": "Karakter rasa tergolong datar",
            "behind_story": "Merek vodka lokal terlaris dari PT Astidama Adimukti."
        },
        {
            "product_name": "Smirnoff Red No. 21",
            "origin_brand": "Diageo (Luar Negeri / Impor)",
            "origin_type": "Luar Negeri / Impor",
            "alcohol_by_volume": "40.0%",
            "base_ingredients": "Gandum impor",
            "price_point": "IDR 300,000 - IDR 450,000",
            "key_difference": "Standar vodka internasional, 10 kali filtrasi",
            "product_weakness": "Harga impor premium",
            "behind_story": "Vodka Rusia yang diproduksi secara global."
        }
    ],
    "wine": [
        {
            "product_name": "Sababay Reserve Red",
            "origin_brand": "PT Sababay Industry (Lokal / Indonesia)",
            "origin_type": "Lokal / Indonesia",
            "alcohol_by_volume": "13.0%",
            "base_ingredients": "Anggur Alphonse Lavallee lokal",
            "price_point": "IDR 250,000 - IDR 350,000",
            "key_difference": "Wine lokal Bali dengan kualitas internasional",
            "product_weakness": "Terbatas pada varietas iklim tropis",
            "behind_story": "Diproduksi di Buleleng, Bali."
        },
        {
            "product_name": "Anggur Kawa-Kawa",
            "origin_brand": "PT Balaraja Barat Indah (Lokal / Indonesia)",
            "origin_type": "Lokal / Indonesia",
            "alcohol_by_volume": "19.8%",
            "base_ingredients": "Fermentasi anggur dan perisa",
            "price_point": "IDR 65,000 - IDR 85,000",
            "key_difference": "Anggur merah komersial dengan ABV sangat tinggi",
            "product_weakness": "Kandungan gula sangat tinggi",
            "behind_story": "Merek unggulan pasar lokal yang sangat terjangkau."
        }
    ],
    "beer": [
        {
            "product_name": "Bir Bintang",
            "origin_brand": "PT Multi Bintang Indonesia Tbk",
            "origin_type": "Lokal / Indonesia",
            "alcohol_by_volume": "4.7%",
            "base_ingredients": "Barley malt, hops, air",
            "price_point": "IDR 25,000 - IDR 40,000",
            "key_difference": "Bir paling ikonik dengan rasa seimbang",
            "product_weakness": "Hanya fokus di varian pilsner & radler",
            "behind_story": "Bir populer milik Heineken Group di Indonesia."
        },
        {
            "product_name": "Anker Beer",
            "origin_brand": "PT Delta Djakarta Tbk",
            "origin_type": "Lokal / Indonesia",
            "alcohol_by_volume": "4.5%",
            "base_ingredients": "Malt, hops unggulan",
            "price_point": "IDR 22,000 - IDR 35,000",
            "key_difference": "Karakter malt lebih kuat",
            "product_weakness": "Market share di bawah kompetitornya",
            "behind_story": "Bir klasik yang diproduksi sejak 1932."
        }
    ],
    "soju": [
        {
            "product_name": "Wija Soju",
            "origin_brand": "PT Balaraja Barat Indah",
            "origin_type": "Lokal / Indonesia",
            "alcohol_by_volume": "15.0%",
            "base_ingredients": "Spirit netral, perisa buah",
            "price_point": "IDR 70,000 - IDR 90,000",
            "key_difference": "Soju lokal dengan harga ekonomis",
            "product_weakness": "Rasa buah cenderung sintetis",
            "behind_story": "Soju lokal untuk pasar hallyu Indonesia."
        },
        {
            "product_name": "Jinro Chamisul Fresh",
            "origin_brand": "HiteJinro (Luar Negeri / Impor)",
            "origin_type": "Luar Negeri / Impor",
            "alcohol_by_volume": "16.5%",
            "base_ingredients": "Beras, barley, tapioka",
            "price_point": "IDR 120,000 - IDR 160,000",
            "key_difference": "Soju No.1 Korea dengan filtrasi bambu",
            "product_weakness": "Harga impor lebih tinggi",
            "behind_story": "Merek spirit paling laris di dunia."
        }
    ],
    "gin": [
        {
            "product_name": "Empire Gin",
            "origin_brand": "PT Astidama Adimukti",
            "origin_type": "Lokal / Indonesia",
            "alcohol_by_volume": "40.0%",
            "base_ingredients": "Spirit lokal, ekstrak juniper berry",
            "price_point": "IDR 140,000 - IDR 180,000",
            "key_difference": "Gin lokal terjangkau untuk mixing",
            "product_weakness": "Aroma kurang kompleks",
            "behind_story": "Gin andalan pasar lokal buatan Astidama."
        },
        {
            "product_name": "Gordon's London Dry Gin",
            "origin_brand": "Diageo (Luar Negeri / Impor)",
            "origin_type": "Luar Negeri / Impor",
            "alcohol_by_volume": "37.5%",
            "base_ingredients": "Juniper, ketumbar",
            "price_point": "IDR 350,000 - IDR 500,000",
            "key_difference": "Rasa juniper tajam & klasik",
            "product_weakness": "Harga impor premium",
            "behind_story": "Merek London Dry Gin legendaris."
        }
    ],
    "rum": [
        {
            "product_name": "Mansion House Rum",
            "origin_brand": "PT Astidama Adimukti",
            "origin_type": "Lokal / Indonesia",
            "alcohol_by_volume": "40.0%",
            "base_ingredients": "Molase tebu",
            "price_point": "IDR 150,000 - IDR 200,000",
            "key_difference": "Rum lokal ekonomis untuk cocktail",
            "product_weakness": "Kurang cocok diminum neat",
            "behind_story": "Rum komersial dari Astidama."
        },
        {
            "product_name": "Bacardi Carta Blanca",
            "origin_brand": "Bacardi Limited (Luar Negeri / Impor)",
            "origin_type": "Luar Negeri / Impor",
            "alcohol_by_volume": "40.0%",
            "base_ingredients": "Molase Karibia",
            "price_point": "IDR 400,000 - IDR 550,000",
            "key_difference": "Rum putih bersih unggulan",
            "product_weakness": "Aroma terlalu ringan",
            "behind_story": "Pionir rum putih global dari Puerto Rico."
        }
    ]
}


# ==============================================================================
# 2. HARD-LOCK SANITIZER (PENGUNCI 2 KOMPETITOR & ANTI HALUSINASI)
# ==============================================================================
def HARD_LOCK_SANITIZER(data_json: dict) -> dict:
    beverages = data_json.get("selected_beverages", [])

    for bev in beverages:
        bev_name = bev.get("beverage_name", "")

        # A. Kunci Fakta Produk Utama
        for key, truth in MASTER_VERIFIED_BRANDS.items():
            if key in bev_name.lower():
                bev["beverage_name"] = truth["real_name"]
                bev["origin_type"] = truth["origin_type"]
                bev["alcohol_by_volume"] = truth["abv"]
                bev["behind_story"] = truth["behind_story"]
                break

        # B. Kunci Fakta Kompetitor
        comps = bev.get("equivalent_competitors", [])
        sanitized_comps = []
        seen_names = set()

        for comp in comps:
            comp_name = comp.get("product_name", "")
            
            # Jika AI halusinasi nama aneh, buang dari list
            if "timantti" in comp_name.lower() or "nusantara" in comp_name.lower() or "fiktif" in comp_name.lower():
                continue

            matched = False
            for key, truth in MASTER_VERIFIED_BRANDS.items():
                if key in comp_name.lower():
                    comp["product_name"] = truth["real_name"]
                    comp["origin_brand"] = f"{truth['producer']} ({truth['origin_type']})"
                    comp["origin_type"] = truth["origin_type"]
                    comp["alcohol_by_volume"] = truth["abv"]
                    comp["behind_story"] = truth["behind_story"]
                    matched = True
                    break
            
            if comp["product_name"] not in seen_names:
                sanitized_comps.append(comp)
                seen_names.add(comp["product_name"])

        # C. GARANSI MINIMAL 2 KOMPETITOR AKTIF 
        # Jika hasil komparasi kurang dari 2, sikat pakai fallback database kita
        category = bev.get("beverage_category", "").lower()
        
        if "alco" in category or "rtd" in category or "pop" in category:
            cat_key = "alco pops"
        elif "vodka" in category:
            cat_key = "vodka"
        elif "wine" in category or "anggur" in category:
            cat_key = "wine"
        elif "beer" in category or "bir" in category:
            cat_key = "beer"
        elif "soju" in category or "sake" in category:
            cat_key = "soju"
        elif "gin" in category:
            cat_key = "gin"
        elif "rum" in category:
            cat_key = "rum"
        else:
            cat_key = "whisky"
            
        fallbacks = REAL_COMPETITOR_FALLBACKS.get(cat_key, REAL_COMPETITOR_FALLBACKS["whisky"])
        
        # Suntik data asli sampai berjumlah tepat 2
        for fb in fallbacks:
            if len(sanitized_comps) >= 2:
                break
            if fb["product_name"] not in seen_names:
                sanitized_comps.append(dict(fb))
                seen_names.add(fb["product_name"])
        
        # Batasi maksimal tampil 3 jika kepanjangan
        bev["equivalent_competitors"] = sanitized_comps[:3]

    return data_json


# ==========================================
# 3. HELPER FUNGSI
# ==========================================
def clean_text(text: str) -> str:
    if not text:
        return ""
    cleaned = re.sub(r"[^\x00-\x7F]+", "", text)
    return re.sub(r"\s+", " ", cleaned).strip()

def fetch_real_product_image(product_name: str) -> str | None:
    try:
        clean_name = clean_text(product_name.split("(")[0])
        if not clean_name: return None
        search_url = f"https://en.wikipedia.org/w/api.php?action=query&generator=search&gsrsearch={urllib.parse.quote(clean_name)}&gsrlimit=1&prop=pageimages&pithumbsize=600&format=json"
        req = urllib.request.Request(search_url, headers={"User-Agent": "Mozilla/5.0"})
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
    if not url or not url.startswith("http"): return None
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=4) as response:
            return io.BytesIO(response.read())
    except Exception:
        return None


# ==========================================
# 4. SCHEMA PYDANTIC & INSTRUCTION AI
# ==========================================
class CompetitorItem(BaseModel):
    product_name: str = Field(description="Nama produk pembanding real")
    origin_brand: str = Field(description="Asal negara / brand produsen")
    origin_type: str = Field(description="Lokal / Indonesia atau Impor")
    alcohol_by_volume: str = Field(description="Kadar alkohol (% ABV)")
    base_ingredients: str = Field(description="Bahan baku")
    price_point: str = Field(description="Rentang harga")
    key_difference: str = Field(description="Perbedaan utama")
    product_weakness: str = Field(description="Kelemahan produk")
    behind_story: str = Field(description="Sejarah singkat")

class BusinessAnalysis(BaseModel):
    target_demographics: str = Field(description="Target konsumen")
    markup_margin_potential: str = Field(description="Potensi Margin")
    suitable_venue: str = Field(description="Venue terbaik")
    investment_value: str = Field(description="Potensi investasi")

class BeverageItem(BaseModel):
    beverage_name: str = Field(description="Nama minuman")
    beverage_category: str = Field(description="Kategori")
    origin_type: str = Field(description="Asal produk")
    alcohol_by_volume: str = Field(description="Kadar alkohol (% ABV)")
    unique_selling_point: str = Field(description="USP")
    behind_story: str = Field(description="Sejarah singkat")
    product_weakness: str = Field(description="Kelemahan")
    age_or_vintage: str = Field(description="Vintage")
    origin: str = Field(description="Negara & Region")
    base_ingredients: str = Field(description="Bahan baku")
    rating_scores: str = Field(description="Rating")
    estimated_price: str = Field(description="Harga")
    flavor_character: list[str] = Field(description="Notes")
    serving_recommendation: str = Field(description="Cara saji")
    business_intelligence: BusinessAnalysis = Field(description="Bisnis")
    equivalent_competitors: list[CompetitorItem] = Field(description="Kompetitor")
    expert_note: str = Field(description="Catatan pakar")
    image_url: str = Field(default="")

class SelectorResponse(BaseModel):
    executive_summary: str = Field(description="Ringkasan")
    selected_beverages: list[BeverageItem] = Field(description="Daftar minuman")


SYSTEM_INSTRUCTION = """
Kamu adalah Master Sommelier. 
KAMI AKAN MENGHAPUS SEMUA MERK FIKTIF. 
Sebutkan sejarah pendirian, kelemahan produk, komposisi bahan baku, serta WAJIB memberikan 2 kompetitor setaranya secara riil. 
Jika kamu tidak tahu kompetitor lokalnya, gunakan merk internasional yang sangat terkenal di pasaran sebagai pembandingnya.
"""


# ==========================================
# 5. EXCEL & PPT GENERATOR
# ==========================================
def create_excel_dashboard(data_json: dict) -> io.BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Executive Dashboard"
    ws.views.sheetView[0].showGridLines = True

    title_font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    title_fill = PatternFill(start_color="4C1D95", end_color="4C1D95", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="6D28D9", end_color="6D28D9", fill_type="solid")
    sec_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    sec_fill = PatternFill(start_color="3B0764", end_color="3B0764", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="E9D5FF"), right=Side(style="thin", color="E9D5FF"),
        top=Side(style="thin", color="E9D5FF"), bottom=Side(style="thin", color="E9D5FF"),
    )

    ws.merge_cells("A1:T2")
    banner = ws["A1"]
    banner.value = "GLOBAL BEVERAGE INTELLIGENCE DASHBOARD"
    banner.font = title_font
    banner.fill = title_fill
    banner.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A4:T6")
    summary_cell = ws["A4"]
    exec_text = data_json.get("executive_summary", "")
    summary_cell.value = f"📌 EXECUTIVE SUMMARY:\n{exec_text}"
    summary_cell.font = Font(name="Calibri", size=10, italic=True)
    summary_cell.fill = PatternFill(start_color="F3E8FF", end_color="F3E8FF", fill_type="solid")
    summary_cell.alignment = Alignment(vertical="center", horizontal="left", wrap_text=True)

    ws_chart_data = wb.create_sheet(title="ChartData")
    ws_chart_data["A1"], ws_chart_data["B1"] = "Kategori", "Jumlah Produk"

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
    pie.title, pie.width, pie.height = "Proporsi Kategori", 14, 7.5
    ws.add_chart(pie, "A8")

    bar = BarChart()
    bar.type, bar.style = "col", 10
    bar.title, bar.y_axis.title = "Sebaran Produk", "Jumlah Botol"
    bar.add_data(data_ref, titles_from_data=True)
    bar.set_categories(labels_ref)
    bar.width, bar.height = 16, 7.5
    ws.add_chart(bar, "I8")

    start_row = 23
    ws.merge_cells(f"A{start_row-1}:T{start_row-1}")
    sec_title = ws[f"A{start_row-1}"]
    sec_title.value = "📊 DETAILED BEVERAGE ANALYSIS"
    sec_title.font, sec_title.fill = sec_font, sec_fill
    sec_title.alignment = Alignment(horizontal="left", vertical="center")

    headers = [
        "No", "Nama Minuman", "Kategori", "Klasifikasi", "ABV", "Vintage", "Asal",
        "Bahan Baku", "Behind Story", "Rating", "Harga", "USP", "Kelemahan",
        "Tasting Notes", "Cara Saji", "Pembanding", "Margin", "Target", "Venue", "Investasi"
    ]

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=start_row, column=col_idx)
        cell.value, cell.font, cell.fill, cell.border = h, header_font, header_fill, thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for idx, item in enumerate(items, 1):
        curr_row = start_row + idx
        biz = item.get("business_intelligence", {})
        comps = item.get("equivalent_competitors", [])

        comp_str = "\n".join([f"• {c.get('product_name')} ({c.get('origin_type')}) | ABV: {c.get('alcohol_by_volume')} | Story: {c.get('behind_story')}" for c in comps])

        row_values = [
            idx, item.get("beverage_name"), item.get("beverage_category"), item.get("origin_type"),
            item.get("alcohol_by_volume"), item.get("age_or_vintage"), item.get("origin"),
            item.get("base_ingredients"), item.get("behind_story"), item.get("rating_scores"),
            item.get("estimated_price"), item.get("unique_selling_point"), item.get("product_weakness"),
            ", ".join(item.get("flavor_character", [])), item.get("serving_recommendation"),
            comp_str, biz.get("markup_margin_potential"), biz.get("target_demographics"),
            biz.get("suitable_venue"), biz.get("investment_value"),
        ]

        for col_idx, val in enumerate(row_values, 1):
            cell = ws.cell(row=curr_row, column=col_idx)
            cell.value, cell.border = val, thin_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if idx % 2 == 0:
                cell.fill = PatternFill(start_color="FAF5FF", end_color="FAF5FF", fill_type="solid")

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max([len(str(cell.value or "")) for cell in col if cell.row >= start_row] + [0])
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 35)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def create_pptx_presentation(data_json: dict) -> io.BytesIO:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    theme = {"card_bg": RGBColor(15, 23, 42), "border": RGBColor(234, 179, 8), "primary": RGBColor(250, 204, 21), "text": RGBColor(241, 245, 249), "card_header": RGBColor(30, 41, 59)}

    slide1 = prs.slides.add_slide(blank_layout)
    bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = theme["card_bg"]
    tf1 = bg1.text_frame
    p1 = tf1.paragraphs[0]
    p1.text, p1.font.size, p1.font.bold, p1.font.color.rgb = "GLOBAL BEVERAGE INTELLIGENCE", Pt(36), True, theme["primary"]

    for idx, item in enumerate(data_json.get("selected_beverages", []), 1):
        slide_a = prs.slides.add_slide(blank_layout)
        bg_a = slide_a.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg_a.fill.solid()
        bg_a.fill.fore_color.rgb = theme["card_bg"]

        card_info = slide_a.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.3))
        card_info.fill.solid()
        card_info.fill.fore_color.rgb = theme["card_header"]
        tf_info = card_info.text_frame
        tf_info.word_wrap = True

        p_spec = tf_info.paragraphs[0]
        p_spec.text = f"#{idx} {item.get('beverage_name')}\n\n• ABV: {item.get('alcohol_by_volume')} | Klasifikasi: {item.get('origin_type')}\n• Behind Story: {item.get('behind_story')}\n• USP: {item.get('unique_selling_point')}"
        p_spec.font.size, p_spec.font.color.rgb = Pt(14), theme["text"]

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer


# ==========================================
# 6. STREAMLIT APP CONFIG & FORM
# ==========================================
st.set_page_config(page_title="Global Beverage & Business AI Pro", page_icon="🥃", layout="wide")
st.title("🍾 Global Beverage Business")
st.write("Beverages Analysis, Sorting & Equivalent Product Benchmarking System.")

def get_auto_api_key() -> str:
    try:
        if "GEMINI_API_KEY" in st.secrets: return st.secrets["GEMINI_API_KEY"]
    except Exception: pass
    return os.environ.get("GEMINI_API_KEY", "")

auto_api_key = get_auto_api_key()

if auto_api_key:
    st.sidebar.success("✅ API Key Terdeteksi Otomatis!")
    raw_api_key = st.sidebar.text_input("Gemini API Key (Opsional / Override):", type="password", value=auto_api_key)
else:
    st.sidebar.warning("⚠️ API Key tidak terdeteksi di .env / Secrets")
    raw_api_key = st.sidebar.text_input("Masukkan Gemini API Key:", type="password", value="")

api_key = raw_api_key.strip()

with st.form("form_analisis"):
    col1, col2 = st.columns(2)
    with col1:
        beverage_type = st.selectbox("🥃 Beverages Category:", ["🌐 Universal / Bebas (Campuran Terbaik)", "🍹 Alco Pops / Ready To Drink (RTD)", "🍷 Wine", "🥃 Whisky / Whiskey", "🍺 Beer / Craft Beer", "🍶 Sake", "🇰🇷 Soju", "🍸 Gin", "🧊 Vodka", "🌵 Tequila / Mezcal", "🏴‍☠️ Rum", "☠ Brandy"])
        origin_scope = st.selectbox("🇲🇨 Scope Of Origin:", ["🌐 Bebas / Campuran (Lokal & Impor)", "🇮🇩 Khusus Lokal / Indonesia", "🌍 Khusus Impor / Luar Negeri", "⚔️ Komparasi Head-to-Head (Lokal vs Impor)"])
        brand_input = st.text_input("🏷️ Brand / Product (Optional):", placeholder="Contoh: Mix Max / Smirnoff Ice / Kura Kura Beer")
        item_count = st.selectbox("🔢 Sort By Range:", options=list(range(1, 16)), index=2)
    with col2:
        sort_priority = st.selectbox("🎯 Main Priority:", ["🌐 Universal / Seimbang", "⭐ Rating Global Tertinggi", "💰 Margin Profit HORECA Tertinggi", "⚖️ Best Value"])
        venue_type = st.selectbox("🏢 Type Of Venue:", ["🌐 Universal / Semua Tipe", "🍸 Speakeasy Bar", "🍷 Fine Dining", "🍻 Casual Pub", "🪩 Nightclub"])
        budget_category = st.selectbox("💵 Price Category:", ["🌐 Universal / Semua Range", "🟢 Entry Level (< $30)", "🟡 Premium Pour ($30 - $100)", "🟠 Top Shelf ($100 - $300)", "🔴 Collector (> $300)"])
        origin_input = st.text_input("🌍 Specific Of Origin (Optional):")

    btn_process = st.form_submit_button("🔍 Punch Down", type="primary")


# ==========================================
# 7. EXECUTION & DISPLAY
# ==========================================
if btn_process:
    if not api_key:
        st.error("❌ API Key tidak terdeteksi. Silakan atur Secrets atau isi di sidebar!")
    else:
        c_bev, c_scope, c_prio, c_ven, c_bud = clean_text(beverage_type), clean_text(origin_scope), clean_text(sort_priority), clean_text(venue_type), clean_text(budget_category)

        prompt_query = f"Sortir {item_count} minuman: Kategori: {c_bev}, Scope Asal: {c_scope}, Brand/Acuan: {brand_input or 'Bebas'}, Prioritas: {c_prio}, Venue: {c_ven}, Budget: {c_bud}, Region: {origin_input or 'Bebas'}. Gunakan hanya merk riil."

        with st.spinner("📊 Menganalisis & menyusun database faktual (Mencegah Halusinasi)..."):
            try:
                client = genai.Client(api_key=api_key)
                try:
                    res = client.models.generate_content(
                        model="gemini-3.6-flash", contents=prompt_query,
                        config=types.GenerateContentConfig(system_instruction=SYSTEM_INSTRUCTION, temperature=0.1, response_mime_type="application/json", response_schema=SelectorResponse)
                    )
                except Exception:
                    res = client.models.generate_content(
                        model="gemini-3.5-flash-lite", contents=prompt_query,
                        config=types.GenerateContentConfig(system_instruction=SYSTEM_INSTRUCTION, temperature=0.1, response_mime_type="application/json", response_schema=SelectorResponse)
                    )

                raw_data = json.loads(res.text)

                # ======================================================
                # EKSEKUSI PEMBASMI HALUSINASI & PEMAKSA 2 KOMPETITOR
                # ======================================================
                data = HARD_LOCK_SANITIZER(raw_data)

                st.success(f"✅ Analisis & Benchmarking {len(data.get('selected_beverages', []))} Minuman Selesai!")

                excel_buf, pptx_buf = create_excel_dashboard(data), create_pptx_presentation(data)

                st.subheader("📥 Export Hasil Laporan")
                col_b1, col_b2 = st.columns(2)
                col_b1.download_button("📊 Download Excel Dashboard (.xlsx)", data=excel_buf, file_name="Executive_Beverage_Dashboard.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
                col_b2.download_button("🖥️ Download PowerPoint (.pptx)", data=pptx_buf, file_name="Executive_Beverage_Presentation.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary", use_container_width=True)

                st.markdown("---")
                st.subheader("📋 Ringkasan Laporan Executive")
                st.info(data["executive_summary"])
                st.markdown("---")

                for i, bev in enumerate(data["selected_beverages"], 1):
                    biz = bev["business_intelligence"]
                    origin_tag = "🇮🇩 LOKAL" if "lokal" in bev.get("origin_type", "").lower() or "indonesia" in bev.get("origin_type", "").lower() else "🌍 IMPOR"

                    st.markdown(f"### #{i} {bev['beverage_name']} ({bev['age_or_vintage']}) &nbsp; `{origin_tag}` &nbsp; 🧪 `{bev['alcohol_by_volume']}`")
                    c1, c2, c3 = st.columns([1.5, 2, 2])
                    
                    c1.markdown(f"**🥃 Kategori:** {bev['beverage_category']}\n\n**🏷️ Klasifikasi:** `{bev['origin_type']}`\n\n**🧪 Kadar Alkohol:** `{bev['alcohol_by_volume']}`\n\n**🌾 Bahan Baku:** {bev['base_ingredients']}\n\n**📍 Asal:** {bev['origin']}\n\n**⭐ Rating:** `{bev['rating_scores']}`\n\n**💵 Harga:** `{bev['estimated_price']}`")
                    c2.markdown(f"**📖 Behind Story:** {bev['behind_story']}\n\n**🔥 USP:** {bev['unique_selling_point']}\n\n**⚠️ Kelemahan:** {bev['product_weakness']}\n\n**👃 Notes:** {', '.join(bev['flavor_character'])}\n\n**🧊 Saji:** {bev['serving_recommendation']}")
                    c3.markdown(f"**📈 Margin:** {biz['markup_margin_potential']}\n\n**👥 Target:** {biz['target_demographics']}\n\n**🏢 Venue:** {biz['suitable_venue']}")

                    st.markdown("#### ⚔️ Produk Pembanding & Benchmarking Setara")
                    comps = bev.get("equivalent_competitors", [])
                    if comps:
                        comp_cols = st.columns(len(comps))
                        for idx_c, comp in enumerate(comps):
                            c_tag = "🇮🇩 LOKAL" if "lokal" in comp.get("origin_type", "").lower() or "indonesia" in comp.get("origin_type", "").lower() else "🌍 IMPOR"
                            comp_cols[idx_c].info(f"**{comp['product_name']}** `{c_tag}`\n\n📖 **Behind Story:** {comp['behind_story']}\n\n🧪 **ABV:** `{comp['alcohol_by_volume']}`\n\n📍 **Asal/Brand:** {comp['origin_brand']}\n\n🌾 **Bahan Baku:** {comp['base_ingredients']}\n\n💵 **Harga:** `{comp['price_point']}`\n\n💡 **Bedanya:** {comp['key_difference']}\n\n⚠️ **Kelemahan:** {comp['product_weakness']}")
                    st.markdown("---")

            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")
