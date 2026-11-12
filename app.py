import re
import sys
import requests
import pandas as pd
import numpy as np
from bs4 import BeautifulSoup

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


def extract_screener_data(url: str) -> dict:
    resp = requests.get(url, headers=HEADERS, timeout=20)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    company_name = (
        soup.find("h1").get_text(strip=True)
        if soup.find("h1")
        else "Company"
    )

    ratios = {}
    for li in soup.select("#top-ratios li"):
        name_el = li.select_one(".name")
        val_el = li.select_one(".number")
        if name_el and val_el:
            ratios[name_el.get_text(strip=True)] = clean_num(
                val_el.get_text(strip=True)
            )

    cmp = ratios.get("Current Price", 0.0)
    face_value = ratios.get("Face Value", 1.0) or 1.0
    div_yield = ratios.get("Dividend Yield", 0.0)
    annual_div = cmp * (div_yield / 100.0)


    def parse_section_table(section_id: str) -> pd.DataFrame:
        sec = soup.find("section", {"id": section_id})
        if not sec:
            raise ValueError(f"Section '{section_id}' not found on page.")
        table = sec.find("table", class_="data-table")
        headers = [th.get_text(strip=True) for th in table.select("thead th")]
        rows = {}
        for tr in table.select("tbody tr"):
            cells = [td.get_text(strip=True) for td in tr.find_all("td")]
            if not cells:
                continue
            metric = re.sub(r"\+$", "", cells[0]).strip()
            rows[metric] = [clean_num(c) for c in cells[1:]]
        df = pd.DataFrame(rows, index=headers[1:])
        df = df[~df.index.str.contains("TTM", case=False, na=False)]
        return df.tail(10)


    pnl_df = parse_section_table("profit-loss")
    bs_df = parse_section_table("balance-sheet")

    common_years = pnl_df.index.intersection(bs_df.index)[-10:]
    pnl_10 = pnl_df.loc[common_years]
    bs_10 = bs_df.loc[common_years]

    return {
        "company": company_name,
        "cmp": cmp,
        "face_value": face_value,
        "annual_dividend": annual_div,
        "years": list(common_years),
        "eps": pnl_10["EPS in Rs"].values,
        "equity_capital": bs_10["Equity Capital"].values,
        "reserves": bs_10["Reserves"].values,
        "borrowings": bs_10["Borrowings"].values,
    }


def run_valuation_model(data: dict) -> dict:
    n = len(data["years"])
    eps = np.array(data["eps"], dtype=float)
    eq_cap = np.array(data["equity_capital"], dtype=float)
    reserves = np.array(data["reserves"], dtype=float)
    borrowings = np.array(data["borrowings"], dtype=float)
    fv = data["face_value"]
    cmp = data["cmp"]


    yoy_growth = [
        ((eps[i + 1] / eps[i]) - 1.0) * 100.0 if eps[i] != 0 else 0.0
        for i in range(n - 1)
    ]
    yoy_growth.append(0.0)


    shares = eq_cap / fv
    gross_bv_ps = (eq_cap + reserves) / shares
    debt_ps = borrowings / shares
    net_bv_ps = gross_bv_ps - debt_ps


    avg_eps = np.mean(eps)
    avg_growth_g = np.mean(yoy_growth[:-1])
    eps_cagr = (
        ((eps[-1] / eps[0]) ** (1.0 / n) - 1.0) * 100.0 if eps[0] > 0 else 0.0
    )


    current_pe = cmp / eps[-1] if eps[-1] > 0 else 0.0
    avg_pe = cmp / avg_eps if avg_eps > 0 else 0.0
    latest_debt_ps = debt_ps[-1]


    iv_7_gross = avg_eps * (7.0 + 1.0 * avg_growth_g)
    iv_85_gross = avg_eps * (8.5 + 2.0 * avg_growth_g)

    iv_7_net = iv_7_gross - latest_debt_ps
    iv_85_net = iv_85_gross - latest_debt_ps

    iv_7_zero_g = (avg_eps * 7.0) - latest_debt_ps
    iv_85_zero_g = (avg_eps * 8.5) - latest_debt_ps


    future_eps_10y = eps[-1] * ((1.0 + eps_cagr / 100.0) ** 10)
    future_price_10y = future_eps_10y * current_pe
    future_div_10y = data["annual_dividend"] * 10.0
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
        "Year": data["years"],
        "EPS (Rs)": np.round(eps, 2),
        "YoY Growth (%)": np.round(yoy_growth, 2),
        "Face Value": fv,
        "Equity Capital (Cr)": eq_cap,
        "Reserves (Cr)": reserves,
        "Borrowings (Cr)": borrowings,
        "Gross BV/Sh": np.round(gross_bv_ps, 2),
        "Debt/Sh": np.round(debt_ps, 2),
        "Net BV/Sh": np.round(net_bv_ps, 2),
    })

    summary = {
        "Company": data["company"],
        "CMP (Rs)": round(cmp, 2),
        "10Y Avg EPS (Rs)": round(avg_eps, 2),
        "Avg Annual EPS Growth G (%)": round(avg_growth_g, 2),
        "10Y EPS CAGR (%)": round(eps_cagr, 2),
        "Current P/E": round(current_pe, 2),
        "Latest Debt/Share (Rs)": round(latest_debt_ps, 2),
        "Graham IV (PE 7, Net of Debt)": round(iv_7_net, 2),
        "Graham IV (PE 8.5, Net of Debt)": round(iv_85_net, 2),
        "Zero-Growth IV (PE 7, Net)": round(iv_7_zero_g, 2),
        "Zero-Growth IV (PE 8.5, Net)": round(iv_85_zero_g, 2),
        "10Y Projected EPS (Rs)": round(future_eps_10y, 2),
        "10Y Projected Price Ex-Div (Rs)": round(future_price_10y, 2),
        "10Y Projected Price + Div (Rs)": round(future_total_val, 2),
        "Projected Price CAGR (%)": round(price_cagr, 2),
        "Projected Total Return CAGR (%)": round(total_cagr, 2),
    }

    return {"historical_table": table_df, "summary": summary}


def export_to_excel(result: dict, filename: str = "Graham_Valuation_Output.xlsx"):
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        result["historical_table"].to_excel(
            writer, sheet_name="10Y_Financials", index=False
        )
        summary_df = pd.DataFrame(
            list(result["summary"].items()), columns=["Metric", "Value"]
        )
        summary_df.to_excel(
            writer, sheet_name="Intrinsic_&_Projection", index=False
        )
    print(f"Exported complete valuation workbook to: {filename}")


if __name__ == "__main__":
    target_url = (
        sys.argv[1]
        if len(sys.argv) > 1
        else "https://www.screener.in/company/TCS/consolidated/"
    )
    raw_data = extract_screener_data(target_url)
    output = run_valuation_model(raw_data)

    print("\n=== 10-YEAR HISTORICAL DATA ===")
    print(output["historical_table"].to_string(index=False))

    print("\n=== GRAHAM INTRINSIC VALUE & 10-YEAR PROJECTION ===")
    for k, v in output["summary"].items():
        print(f"{k:<35}: {v}")

    export_to_excel(output)