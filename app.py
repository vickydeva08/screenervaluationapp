from io import BytesIO
import re
from bs4 import BeautifulSoup
from fpdf import FPDF
import numpy as np
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
import requests
import streamlit as st

st.set_page_config(
    page_title="Stock Valuation & Presentation Report",
    page_icon="📊",
    layout="wide",
)

# Presentation-style CSS
st.markdown(
    """
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .report-header {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        color: white;
        padding: 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .report-header h1 {
        margin: 0;
        font-size: 2rem;
        font-weight: 700;
    }
    .report-header p {
        margin: 0.5rem 0 0 0;
        opacity: 0.9;
        font-size: 1rem;
    }
    .metric-card {
        background: white;
        padding: 1.2rem;
        border-radius: 10px;
        border-left: 5px solid #2563eb;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 1rem;
    }
    .metric-label {
        font-size: 0.85rem;
        color: #64748b;
        font-weight: 600;
        text-transform: uppercase;
    }
    .metric-value {
        font-size: 1.5rem;
        color: #0f172a;
        font-weight: 700;
        margin-top: 0.25rem;
    }
    .section-title {
        color: #1e293b;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 0.5rem;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
        font-weight: 700;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="report-header">
        <h1>📊 Equity Valuation & Fundamental Report</h1>
        <p>Automated 10-Year Graham Intrinsic Value (PE 7 & PE 8.5), Return Ratios & Future Projection Generator</p>
    </div>
    """,
    unsafe_allow_html=True,
)

url = st.text_input(
    "Enter Screener.in Consolidated/Standalone URL:",
    "https://www.screener.in/company/TCS/consolidated/",
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}


def clean_num(val: str) -> float:
  if not val:
    return 0.0
  cleaned = re.sub(r"[^\d.\-]", "", str(val))
  try:
    return float(cleaned)
  except ValueError:
    return 0.0


def render_card(label: str, value: str):
  st.markdown(
      f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
        </div>
        """,
      unsafe_allow_html=True,
  )


@st.cache_data
def extract_screener_data(url: str) -> dict:
  resp = requests.get(url, headers=HEADERS, timeout=20)
  resp.raise_for_status()
  soup = BeautifulSoup(resp.text, "html.parser")

  company_name = (
      soup.find("h1").get_text(strip=True) if soup.find("h1") else "Company"
  )

  ratios = {}
  for li in soup.select("#top-ratios li"):
    name_el = li.select_one(".name")
    val_el = li.select_one(".number")
    if name_el and val_el:
      ratios[name_el.get_text(strip=True)] = clean_num(
          val_el.get_text(strip=True)
      )

  market_cap = ratios.get("Market Cap", 0.0)
  cmp = ratios.get("Current Price", 0.0)
  stock_pe = ratios.get("Stock P/E", 0.0)
  face_value = ratios.get("Face Value", 1.0) or 1.0
  div_yield = ratios.get("Dividend Yield", 0.0)
  annual_div = cmp * (div_yield / 100.0)
  top_roce = ratios.get("ROCE", 0.0)
  top_roe = ratios.get("ROE", 0.0)

  def parse_section_table(section_id: str) -> pd.DataFrame:
    sec = soup.find("section", {"id": section_id})
    if not sec:
      return pd.DataFrame()
    table = sec.find("table", class_="data-table")
    if not table:
      return pd.DataFrame()
    headers = [th.get_text(strip=True) for th in table.select("thead th")]
    rows = {}
    for tr in table.select("tbody tr"):
      cells = [td.get_text(strip=True) for td in tr.find_all("td")]
      if not cells:
        continue
      metric = re.sub(r"\+$", "", cells[0]).strip()
      rows[metric] = [clean_num(c) for c in cells[1:]]
    df = pd.DataFrame(rows, index=headers[1:])
    return df

  pnl_raw = parse_section_table("profit-loss")
  bs_raw = parse_section_table("balance-sheet")
  ratios_raw = parse_section_table("ratios")

  if pnl_raw.empty or bs_raw.empty:
    raise ValueError("Required Profit & Loss or Balance Sheet tables not found.")

  ttm_rows = pnl_raw[pnl_raw.index.str.contains("TTM", case=False, na=False)]
  ttm_eps = (
      float(ttm_rows["EPS in Rs"].iloc[-1])
      if (not ttm_rows.empty and "EPS in Rs" in ttm_rows.columns)
      else 0.0
  )

  pnl_df = pnl_raw[~pnl_raw.index.str.contains("TTM", case=False, na=False)].tail(10)
  bs_df = bs_raw[~bs_raw.index.str.contains("TTM", case=False, na=False)].tail(10)
  ratios_df = ratios_raw[~ratios_raw.index.str.contains("TTM", case=False, na=False)].tail(10)

  common_years = pnl_df.index.intersection(bs_df.index)[-10:]
  pnl_10 = pnl_df.loc[common_years]
  bs_10 = bs_df.loc[common_years]

  roce_list = []
  roe_list = []
  for yr in common_years:
    if not ratios_df.empty and yr in ratios_df.index and "ROCE %" in ratios_df.columns:
      roce_list.append(ratios_df.loc[yr, "ROCE %"])
    else:
      roce_list.append(0.0)

    if not ratios_df.empty and yr in ratios_df.index and "ROE %" in ratios_df.columns:
      roe_list.append(ratios_df.loc[yr, "ROE %"])
    else:
      np_val = pnl_10.loc[yr, "Net Profit"] if "Net Profit" in pnl_10.columns else 0.0
      nw_val = bs_10.loc[yr, "Equity Capital"] + bs_10.loc[yr, "Reserves"]
      roe_list.append((np_val / nw_val) * 100.0 if nw_val > 0 else 0.0)

  return {
      "company": company_name,
      "market_cap": market_cap,
      "cmp": cmp,
      "stock_pe": stock_pe,
      "ttm_eps": ttm_eps,
      "face_value": face_value,
      "annual_dividend": annual_div,
      "top_roce": top_roce,
      "top_roe": top_roe,
      "years": list(common_years),
      "eps": pnl_10["EPS in Rs"].values,
      "equity_capital": bs_10["Equity Capital"].values,
      "reserves": bs_10["Reserves"].values,
      "borrowings": bs_10["Borrowings"].values,
      "roce": np.array(roce_list, dtype=float),
      "roe": np.array(roe_list, dtype=float),
  }


def generate_pdf(summary: dict, table_df: pd.DataFrame) -> bytes:
  pdf = FPDF(orientation="L", unit="mm", format="A4")
  pdf.add_page()

  pdf.set_fill_color(30, 58, 138)
  pdf.rect(0, 0, 297, 25, "F")
  pdf.set_text_color(255, 255, 255)
  pdf.set_font("Helvetica", "B", 16)
  pdf.set_xy(10, 8)
  pdf.cell(0, 10, f"{summary['company']} - Fundamental & Valuation Report", ln=True)

  pdf.set_text_color(0, 0, 0)
  pdf.set_xy(10, 32)
  pdf.set_font("Helvetica", "B", 11)
  pdf.cell(0, 8, "Key Valuation & Return Metrics", ln=True)

  pdf.set_font("Helvetica", "", 10)
  metrics_items = [
      ("Market Cap", f"Rs. {summary['market_cap']:,.2f} Cr."),
      ("Current Price (CMP)", f"Rs. {summary['cmp']:.2f}"),
      ("Annual Dividend", f"Rs. {summary['annual_dividend']:.2f}"),
      ("Current P/E", f"{summary['current_pe']:.2f}x"),
      ("10Y Average P/E", f"{summary['avg_pe']:.2f}x"),
      ("10Y EPS CAGR", f"{summary['eps_cagr']:.2f}%"),
      ("10Y Avg ROE", f"{summary['avg_roe']:.2f}%"),
      ("10Y Avg ROCE", f"{summary['avg_roce']:.2f}%"),
      ("Graham IV (PE 7, Net)", f"Rs. {summary['iv_7_net']:,.2f}"),
      ("Graham IV (PE 8.5, Net)", f"Rs. {summary['iv_85_net']:,.2f}"),
      ("Projected 10Y Price", f"Rs. {summary['future_price_10y']:,.2f}"),
      ("Projected 10Y Total Value", f"Rs. {summary['future_total_val']:,.2f}"),
  ]

  col_w = 68
  row_h = 8
  start_y = 42
  for idx, (k, v) in enumerate(metrics_items):
    r = idx // 4
    c = idx % 4
    pdf.set_xy(10 + c * col_w, start_y + r * row_h)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(42, row_h, f"{k}:", border=1)
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(26, row_h, v, border=1)

  pdf.set_xy(10, 75)
  pdf.set_font("Helvetica", "B", 11)
  pdf.cell(0, 8, "10-Year Historical Performance", ln=True)

  cols = list(table_df.columns)
  t_col_w = 277 / len(cols)
  pdf.set_font("Helvetica", "B", 8)
  pdf.set_fill_color(226, 232, 240)
  for col in cols:
    pdf.cell(t_col_w, 7, str(col), border=1, fill=True, align="C")
  pdf.ln()

  pdf.set_font("Helvetica", "", 8)
  for _, row in table_df.iterrows():
    for col in cols:
      pdf.cell(t_col_w, 6.5, str(row[col]), border=1, align="C")
    pdf.ln()

  return bytes(pdf.output())


def generate_pptx(summary: dict, table_df: pd.DataFrame) -> bytes:
  prs = Presentation()
  prs.slide_width = Inches(13.333)
  prs.slide_height = Inches(7.5)

  slide1 = prs.slides.add_slide(prs.slide_layouts[6])
  title_box = slide1.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12.3), Inches(0.8))
  tf = title_box.text_frame
  p = tf.paragraphs[0]
  p.text = f"{summary['company']} - Equity Valuation & Projections"
  p.font.size = Pt(26)
  p.font.bold = True

  metrics_items = [
      ("Market Cap", f"Rs. {summary['market_cap']:,.2f} Cr."),
      ("Current Price (CMP)", f"Rs. {summary['cmp']:.2f}"),
      ("Annual Dividend", f"Rs. {summary['annual_dividend']:.2f}"),
      ("Current P/E", f"{summary['current_pe']:.2f}x"),
      ("10Y Average P/E", f"{summary['avg_pe']:.2f}x"),
      ("10Y EPS CAGR", f"{summary['eps_cagr']:.2f}%"),
      ("10Y Avg ROE", f"{summary['avg_roe']:.2f}%"),
      ("10Y Avg ROCE", f"{summary['avg_roce']:.2f}%"),
      ("Graham IV (PE 7, Net)", f"Rs. {summary['iv_7_net']:,.2f}"),
      ("Graham IV (PE 8.5, Net)", f"Rs. {summary['iv_85_net']:,.2f}"),
      ("Projected 10Y Price", f"Rs. {summary['future_price_10y']:,.2f}"),
      ("Projected 10Y Dividends", f"Rs. {summary['future_div_10y']:,.2f}"),
      ("Total 10Y Value (Price+Div)", f"Rs. {summary['future_total_val']:,.2f}"),
      ("Projected Price CAGR", f"{summary['price_cagr']:.2f}%"),
      ("Projected Total CAGR", f"{summary['total_cagr']:.2f}%"),
  ]

  rows, cols = 5, 3
  tbl_shape = slide1.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(12.3), Inches(5.2))
  tbl = tbl_shape.table
  for idx, (k, v) in enumerate(metrics_items):
    r = idx // cols
    c = idx % cols
    cell = tbl.cell(r, c)
    cell.text = f"{k}\n{v}"
    for para in cell.text_frame.paragraphs:
      para.font.size = Pt(13)
    cell.text_frame.paragraphs[0].font.bold = True

  slide2 = prs.slides.add_slide(prs.slide_layouts[6])
  title_box2 = slide2.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12.3), Inches(0.8))
  p2 = title_box2.text_frame.paragraphs[0]
  p2.text = f"{summary['company']} - 10-Year Financial & Return Metrics"
  p2.font.size = Pt(24)
  p2.font.bold = True

  n_rows = len(table_df) + 1
  n_cols = len(table_df.columns)
  t_shape2 = slide2.shapes.add_table(n_rows, n_cols, Inches(0.5), Inches(1.4), Inches(12.3), Inches(5.5))
  tbl2 = t_shape2.table

  for c_idx, col_name in enumerate(table_df.columns):
    cell = tbl2.cell(0, c_idx)
    cell.text = str(col_name)
    cell.text_frame.paragraphs[0].font.size = Pt(11)
    cell.text_frame.paragraphs[0].font.bold = True

  for r_idx, (_, row) in enumerate(table_df.iterrows(), start=1):
    for c_idx, col_name in enumerate(table_df.columns):
      cell = tbl2.cell(r_idx, c_idx)
      cell.text = str(row[col_name])
      cell.text_frame.paragraphs[0].font.size = Pt(11)

  buf = BytesIO()
  prs.save(buf)
  return buf.getvalue()


if st.button("Run Valuation & Generate Report", type="primary"):
  try:
    with st.spinner("Fetching 10-year financial data and generating report..."):
      raw = extract_screener_data(url)
      n = len(raw["years"])
      eps = np.array(raw["eps"], dtype=float)
      eq_cap = np.array(raw["equity_capital"], dtype=float)
      reserves = np.array(raw["reserves"], dtype=float)
      borrowings = np.array(raw["borrowings"], dtype=float)
      roce = raw["roce"]
      roe = raw["roe"]
      fv = raw["face_value"]
      cmp = raw["cmp"]
      market_cap = raw["market_cap"]
      annual_div = raw["annual_dividend"]

      yoy_growth = [
          (((eps[i + 1] / eps[i]) - 1.0) * 100.0 if eps[i] != 0 else 0.0)
          for i in range(n - 1)
      ]
      yoy_growth.append(0.0)

      shares = eq_cap / fv
      gross_bv_ps = (eq_cap + reserves) / shares
      debt_ps = borrowings / shares
      net_bv_ps = gross_bv_ps - debt_ps

      avg_eps = float(np.mean(eps))
      avg_roe = float(np.mean(roe)) if len(roe) > 0 else raw["top_roe"]
      avg_roce = float(np.mean(roce)) if len(roce) > 0 else raw["top_roce"]

      cagr_periods = n
      eps_cagr = (
          round(((eps[-1] / eps[0]) ** (1.0 / cagr_periods) - 1.0) * 100.0, 2)
          if eps[0] > 0
          else 0.0
      )

      if raw.get("stock_pe", 0.0) > 0:
        current_pe = raw["stock_pe"]
      elif raw.get("ttm_eps", 0.0) > 0:
        current_pe = cmp / raw["ttm_eps"]
      else:
        current_pe = cmp / eps[-1] if eps[-1] > 0 else 0.0

      avg_pe = cmp / avg_eps if avg_eps > 0 else 0.0
      latest_debt_ps = debt_ps[-1]

      # Graham Intrinsic Value using PE 7 and PE 8.5 with 10Y EPS CAGR
      iv_7_gross = avg_eps * (7.0 + 1.0 * eps_cagr)
      iv_85_gross = avg_eps * (8.5 + 2.0 * eps_cagr)

      iv_7_net = iv_7_gross - latest_debt_ps
      iv_85_net = iv_85_gross - latest_debt_ps

      g = eps_cagr / 100.0
      future_eps_10y = eps[-1] * ((1.0 + g) ** 10)
      future_price_10y = future_eps_10y * current_pe
      future_div_10y = annual_div * 10.0
      future_total_val = future_price_10y + future_div_10y

      price_cagr = (
          ((future_price_10y / cmp) ** (1.0 / 10.0) - 1.0) * 100.0
          if cmp > 0
          else 0.0
      )
      total_cagr = (
          ((future_total_val / cmp) ** (1.0 / 10.0) - 1.0) * 100.0
          if cmp > 0
          else 0.0
      )

      table_df = pd.DataFrame({
          "Year": raw["years"],
          "EPS (Rs)": np.round(eps, 2),
          "YoY Growth (%)": np.round(yoy_growth, 2),
          "ROE (%)": np.round(roe, 2),
          "ROCE (%)": np.round(roce, 2),
          "Gross BV/Sh": np.round(gross_bv_ps, 2),
          "Debt/Sh": np.round(debt_ps, 2),
          "Net BV/Sh": np.round(net_bv_ps, 2),
      })

      summary_data = {
          "company": raw["company"],
          "market_cap": market_cap,
          "cmp": cmp,
          "annual_dividend": annual_div,
          "current_pe": current_pe,
          "avg_pe": avg_pe,
          "eps_cagr": eps_cagr,
          "avg_roe": avg_roe,
          "avg_roce": avg_roce,
          "iv_7_net": iv_7_net,
          "iv_85_net": iv_85_net,
          "future_eps_10y": future_eps_10y,
          "future_price_10y": future_price_10y,
          "future_div_10y": future_div_10y,
          "future_total_val": future_total_val,
          "price_cagr": price_cagr,
          "total_cagr": total_cagr,
      }

      st.success(f"Presentation Report Ready: {raw['company']}")

      dl_col1, dl_col2, _ = st.columns([1, 1, 2])
      with dl_col1:
        pdf_bytes = generate_pdf(summary_data, table_df)
        st.download_button(
            label="📄 Download PDF Report",
            data=pdf_bytes,
            file_name=f"{raw['company']}_Valuation_Report.pdf",
            mime="application/pdf",
        )
      with dl_col2:
        pptx_bytes = generate_pptx(summary_data, table_df)
        st.download_button(
            label="📽️ Download PPTX Presentation",
            data=pptx_bytes,
            file_name=f"{raw['company']}_Valuation_Presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

      st.markdown('<h3 class="section-title">1. Company Overview & Return Ratios</h3>', unsafe_allow_html=True)
      c1, c2, c3, c4 = st.columns(4)
      with c1:
        render_card("Market Capitalization", f"Rs. {market_cap:,.2f} Cr.")
      with c2:
        render_card("Annual Dividend / Share", f"Rs. {annual_div:.2f}")
      with c3:
        render_card("Current Stock Price (CMP)", f"Rs. {cmp:.2f}")
      with c4:
        render_card("10Y EPS CAGR", f"{eps_cagr:.2f}%")

      c5, c6, c7, c8 = st.columns(4)
      with c5:
        render_card("Current P/E Ratio", f"{current_pe:.2f}x")
      with c6:
        render_card("10Y Average P/E Ratio", f"{avg_pe:.2f}x")
      with c7:
        render_card("10Y Average ROE", f"{avg_roe:.2f}%")
      with c8:
        render_card("10Y Average ROCE", f"{avg_roce:.2f}%")

      st.markdown('<h3 class="section-title">2. Intrinsic Value & 10-Year Projections</h3>', unsafe_allow_html=True)
      p1, p2, p3 = st.columns(3)
      with p1:
        render_card("Graham IV (PE 7, Net)", f"Rs. {iv_7_net:,.2f}")
        render_card("Projected 10Y EPS", f"Rs. {future_eps_10y:.2f}")
      with p2:
        render_card("Graham IV (PE 8.5, Net)", f"Rs. {iv_85_net:,.2f}")
        render_card("Projected 10Y Stock Price", f"Rs. {future_price_10y:.2f}")
      with p3:
        render_card("Projected 10Y Dividends", f"Rs. {future_div_10y:.2f}")
        render_card(
            "Total 10Y Value (Price + Div)",
            f"Rs. {future_total_val:,.2f} ({total_cagr:.2f}% CAGR)",
        )

      st.markdown('<h3 class="section-title">3. 10-Year Historical Financial Data</h3>', unsafe_allow_html=True)
      st.dataframe(table_df, use_container_width=True, hide_index=True)

  except Exception as e:
    st.error(f"Error: {e}")