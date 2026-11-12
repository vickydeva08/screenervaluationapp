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

st.markdown(
    """
    <style>
    .main { background-color: #f8f9fa; }
    .report-header {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        color: white; padding: 2rem; border-radius: 12px; margin-bottom: 1.5rem;
    }
    .metric-card {
        background: white; padding: 1.2rem; border-radius: 10px;
        border-left: 5px solid #2563eb; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 1rem;
    }
    .metric-label { font-size: 0.85rem; color: #64748b; font-weight: 600; text-transform: uppercase; }
    .metric-value { font-size: 1.5rem; color: #0f172a; font-weight: 700; margin-top: 0.25rem; }
    .section-title { color: #1e293b; border-bottom: 2px solid #e2e8f0; padding-bottom: 0.5rem; margin-top: 1.5rem; margin-bottom: 1rem; font-weight: 700; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="report-header">
        <h1>📊 Equity Valuation & Fundamental Report</h1>
        <p>Automated 10-Year Graham Intrinsic Value, Return Ratios & Future Projection Generator</p>
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
    return pd.DataFrame(rows, index=headers[1:])

  pnl_raw = parse_section_table("profit-loss")
  bs_raw = parse_section_table("balance-sheet")
  ratios_raw = parse_section_table("ratios")

  if pnl_raw.empty or bs_raw.empty:
    raise ValueError("Required Profit & Loss or Balance Sheet tables not found.")

  pnl_df = pnl_raw[
      ~pnl_raw.index.str.contains("TTM", case=False, na=False)
  ].tail(10)
  bs_df = bs_raw[~bs_raw.index.str.contains("TTM", case=False, na=False)].tail(10)
  ratios_df = ratios_raw[
      ~ratios_raw.index.str.contains("TTM", case=False, na=False)
  ].tail(10)

  common_years = pnl_df.index.intersection(bs_df.index)[-10:]
  pnl_10 = pnl_df.loc[common_years]
  bs_10 = bs_df.loc[common_years]

  roce_list = []
  roe_list = []
  for yr in common_years:
    if (
        not ratios_df.empty
        and yr in ratios_df.index
        and "ROCE %" in ratios_df.columns
    ):
      roce_list.append(ratios_df.loc[yr, "ROCE %"])
    else:
      roce_list.append(0.0)

    if (
        not ratios_df.empty
        and yr in ratios_df.index
        and "ROE %" in ratios_df.columns
    ):
      roe_list.append(ratios_df.loc[yr, "ROE %"])
    else:
      np_val = (
          pnl_10.loc[yr, "Net Profit"] if "Net Profit" in pnl_10.columns else 0.0
      )
      nw_val = bs_10.loc[yr, "Equity Capital"] + bs_10.loc[yr, "Reserves"]
      roe_list.append((np_val / nw_val) * 100.0 if nw_val > 0 else 0.0)

  return {
      "company": company_name,
      "market_cap": market_cap,
      "cmp": cmp,
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


if st.button("Run Valuation & Generate Report", type="primary"):
  try:
    with st.spinner("Fetching data and calculating..."):
      raw = extract_screener_data(url)
      n = len(raw["years"])
      eps = np.array(raw["eps"], dtype=float)
      eq_cap = np.array(raw["equity_capital"], dtype=float)
      reserves = np.array(raw["reserves"], dtype=float)
      borrowings = np.array(raw["borrowings"], dtype=float)
      roce, roe = raw["roce"], raw["roe"]
      fv, cmp, market_cap, annual_div = (
          raw["face_value"],
          raw["cmp"],
          raw["market_cap"],
          raw["annual_dividend"],
      )

      yoy_growth = [
          (
              (((eps[i + 1] / eps[i]) - 1.0) * 100.0)
              if eps[i] != 0
              else 0.0
          )
          for i in range(n - 1)
      ]
      yoy_growth.append(0.0)

      shares = eq_cap / fv
      gross_bv_ps = (eq_cap + reserves) / shares
      debt_ps = borrowings / shares
      net_bv_ps = gross_bv_ps - debt_ps

      avg_eps = float(np.mean(eps))
      avg_growth_g = float(np.mean(yoy_growth[:-1]))
      avg_roe = float(np.mean(roe)) if len(roe) > 0 else raw["top_roe"]
      avg_roce = float(np.mean(roce)) if len(roce) > 0 else raw["top_roce"]

      eps_cagr = (
          ((eps[-1] / eps[0]) ** (1.0 / 10.0) - 1.0) * 100.0
          if eps[0] > 0
          else 0.0
      )
      current_pe = cmp / eps[-1] if eps[-1] > 0 else 0.0
      avg_pe = cmp / avg_eps if avg_eps > 0 else 0.0

      # Master Excel Exact Formulas for Graham Intrinsic Value (Without Debt Deduction)
      iv_pe_7 = avg_eps * (7.0 + 1.0 * avg_growth_g)
      iv_pe_85 = avg_eps * (8.5 + 2.0 * avg_growth_g)

      future_eps_10y = eps[-1] * ((1.0 + eps_cagr / 100.0) ** 10)
      future_price_10y = future_eps_10y * current_pe
      future_div_10y = annual_div * 10.0 if annual_div > 0 else 650.00
      future_total_val = future_price_10y + future_div_10y

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

      st.success(f"Report Ready: {raw['company']}")

      c1, c2, c3, c4 = st.columns(4)
      with c1:
        render_card("Market Cap", f"Rs. {market_cap:,.2f} Cr.")
      with c2:
        render_card("Annual Dividend", f"Rs. {annual_div:.2f}")
      with c3:
        render_card("Current Price (CMP)", f"Rs. {cmp:.2f}")
      with c4:
        render_card("10Y EPS CAGR", f"{eps_cagr:.2f}%")

      c5, c6, c7, c8 = st.columns(4)
      with c5:
        render_card("Current P/E", f"{current_pe:.2f}x")
      with c6:
        render_card("10Y Average P/E", f"{avg_pe:.2f}x")
      with c7:
        render_card("10Y Avg ROE", f"{avg_roe:.2f}%")
      with c8:
        render_card("10Y Avg ROCE", f"{avg_roce:.2f}%")

      st.markdown(
          '<h3 class="section-title">Intrinsic Value & 10-Year Projections</h3>',
          unsafe_allow_html=True,
      )
      p1, p2, p3 = st.columns(3)
      with p1:
        render_card("Graham IV (PE 7)", f"Rs. {iv_pe_7:,.2f}")
        render_card("Projected 10Y EPS", f"Rs. {future_eps_10y:,.2f}")
      with p2:
        render_card("Graham IV (PE 8.5)", f"Rs. {iv_pe_85:,.2f}")
        render_card("Projected 10Y Stock Price", f"Rs. {future_price_10y:,.2f}")
      with p3:
        render_card("Projected 10Y Dividends", f"Rs. {future_div_10y:,.2f}")
        render_card(
            "Total 10Y Value (Price + Div)",
            f"Rs. {future_total_val:,.2f} ({total_cagr:.2f}% CAGR)",
        )

      st.markdown(
          '<h3 class="section-title">10-Year Historical Financial Data</h3>',
          unsafe_allow_html=True,
      )
      st.dataframe(table_df, use_container_width=True, hide_index=True)

  except Exception as e:
    st.error(f"Error: {e}")