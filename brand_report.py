import sys
import os
sys.path.insert(0, r"C:\Users\Usuario\.cursor\skills\pesquisa-marca-dataforseo\scripts")

from client import RestClient
import openpyxl

KEYWORDS = ["bankme", "finscale", "giro tech", "somos valor"]
URLS = {
    "finscale": "https://finscale.com.br/",
    "bankme": "https://bankme.tech/",
    "giro tech": "https://giro.tech/",
    "somos valor": "https://somosvalor.com.br/",
}
FOCUS = "somos valor"

OVERRIDES = {
    "finscale": {"2026-03": 1000},
}

CHANNELS = {
    "meta": {
        "title": "Meta Ads",
        "subtitle": "Anúncios ativos na Biblioteca de Anúncios da Meta (Facebook + Instagram)",
        "metric": "Anúncios ativos",
        "cost_per_ad": 400,
        "data": {
            "finscale": 49,
            "bankme": 0,
            "giro tech": 0,
            "somos valor": 0,
        },
    },
    "ads": {
        "title": "Google Ads",
        "subtitle": "Anúncios ativos no Google Ads",
        "metric": "Anúncios ativos",
        "cost_per_ad": 200,
        "data": {
            "bankme": 59,
            "finscale": 12,
            "giro tech": 3,
            "somos valor": 0,
        },
    },
    "seo": {
        "title": "SEO Orgânico",
        "subtitle": "Tráfego orgânico mensal estimado",
        "metric": "Visitas orgânicas/mês",
        "cost_per_ad": None,
        "data": {
            "giro tech": 6900,
            "bankme": 1800,
            "finscale": 95,
            "somos valor": 0,
        },
    },
    "instagram": {
        "title": "Instagram",
        "subtitle": "Seguidores no perfil oficial",
        "metric": "Seguidores",
        "cost_per_ad": None,
        "data": {
            "finscale": 29000,
            "bankme": 21000,
            "giro tech": 1400,
            "somos valor": 1400,
        },
    },
}
LOCATION_CODE = 2076  # Brasil
OUTPUT_XLSX = r"c:\Users\Usuario\OneDrive\Documentos\0104 unico\relatorio_marca_somosvalor.xlsx"
OUTPUT_HTML = r"c:\Users\Usuario\OneDrive\Documentos\0104 unico\relatorio_marca_somosvalor.html"


def calc_growth(new, old):
    if old in [None, 0] or new is None:
        return None
    return round(((new - old) / old) * 100, 2)


def fetch(client, keywords):
    post_data = {
        0: {
            "location_code": LOCATION_CODE,
            "keywords": keywords,
            "date_from": "2022-01-01",
            "date_to": "2026-03-31",
            "search_partners": False,
        }
    }
    response = client.post("/v3/keywords_data/google_ads/search_volume/live", post_data)
    if response.get("status_code") != 20000:
        print("Erro:", response.get("status_code"), response.get("status_message"))
        return []
    return response["tasks"][0]["result"] or []


def parse(kw):
    item = {
        "keyword": kw["keyword"],
        "2023-03": None,
        "2024-03": None,
        "2025-03": None,
        "2026-03": None,
    }
    for v in kw.get("monthly_searches") or []:
        if v["month"] != 3:
            continue
        key = f"{v['year']}-03"
        if key in item:
            item[key] = v["search_volume"]
    if kw["keyword"] in OVERRIDES:
        for k, v in OVERRIDES[kw["keyword"]].items():
            item[k] = v
    item["growth_1y"] = calc_growth(item["2026-03"], item["2025-03"])
    item["growth_3y"] = calc_growth(item["2026-03"], item["2023-03"])
    return item


def main():
    client = RestClient("rafaelpiromdi@gmail.com", "6f0c916726277f02")

    raw = fetch(client, KEYWORDS)
    by_kw = {r["keyword"]: r for r in raw}

    for kw in KEYWORDS:
        entry = by_kw.get(kw)
        if not entry or not entry.get("monthly_searches"):
            retry = fetch(client, [kw, kw + " brasil", kw + " app", kw + " login"])
            for r in retry:
                if r["keyword"] == kw and r.get("monthly_searches"):
                    by_kw[kw] = r
                    break

    results = []
    for kw in KEYWORDS:
        entry = by_kw.get(kw)
        if entry:
            results.append(parse(entry))
        else:
            results.append({"keyword": kw, "2023-03": None, "2024-03": None, "2025-03": None, "2026-03": None, "growth_1y": None, "growth_3y": None})

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Relatorio Marca"
    headers = [
        "Marca",
        "Crescimento 1 ano (%)",
        "Crescimento 3 anos (%)",
        "03/2026",
        "03/2025",
        "03/2024",
        "03/2023",
    ]
    ws.append(headers)
    for r in results:
        ws.append([
            r["keyword"],
            r["growth_1y"],
            r["growth_3y"],
            r["2026-03"],
            r["2025-03"],
            r["2024-03"],
            r["2023-03"],
        ])

    for col_idx, h in enumerate(headers, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = max(14, len(h) + 2)

    saved_xlsx = OUTPUT_XLSX
    f_idx = 1
    while True:
        try:
            wb.save(saved_xlsx)
            break
        except PermissionError:
            base, ext = os.path.splitext(OUTPUT_XLSX)
            saved_xlsx = f"{base}_{f_idx}{ext}"
            f_idx += 1
    print("OK ->", saved_xlsx)

    saved_html = OUTPUT_HTML
    h_idx = 1
    while True:
        try:
            write_html(results, saved_html)
            break
        except PermissionError:
            base, ext = os.path.splitext(OUTPUT_HTML)
            saved_html = f"{base}_{h_idx}{ext}"
            h_idx += 1
    print("OK ->", saved_html)

    print()
    print("Resumo:")
    for r in results:
        print(f"  {r['keyword']:<14}  03/26={r['2026-03']}  03/25={r['2025-03']}  03/24={r['2024-03']}  03/23={r['2023-03']}  1y={r['growth_1y']}%  3y={r['growth_3y']}%")


def fmt_int(v):
    if v is None:
        return "&mdash;"
    return f"{v:,}".replace(",", ".")


def fmt_brl(v):
    if v is None:
        return "&mdash;"
    return "R$ " + f"{v:,.0f}".replace(",", ".")


def render_channel(channel_key, channel):
    rows = sorted(channel["data"].items(), key=lambda kv: kv[1], reverse=True)
    has_cost = channel["cost_per_ad"] is not None
    cost_per_ad = channel["cost_per_ad"]
    max_val = max(channel["data"].values()) or 1

    body_rows = []
    for kw, val in rows:
        is_focus = kw == FOCUS
        url = URLS.get(kw, "#")
        row_class = ' class="focus"' if is_focus else ""
        bar_pct = (val / max_val * 100) if max_val else 0
        invest = val * cost_per_ad if has_cost else None

        invest_cell = f'<td class="num invest">{fmt_brl(invest)}</td>' if has_cost else ""

        body_rows.append(f"""
            <tr{row_class}>
                <td class="brand">
                    <div class="brand-name">{kw.upper()}{' <span class="focus-badge">FOCO</span>' if is_focus else ''}</div>
                    <a class="brand-url" href="{url}" target="_blank" rel="noopener">{url}</a>
                </td>
                <td class="num highlight">{fmt_int(val)}</td>
                <td class="bar-cell"><div class="bar-track"><div class="bar-fill" style="width:{bar_pct:.1f}%"></div></div></td>
                {invest_cell}
            </tr>
        """.strip())

    invest_header = '<th class="num">Estimativa Investimento</th>' if has_cost else ""
    cost_legend = (
        f'<span class="meta-tag">R$ {cost_per_ad} por anúncio</span>'
        if has_cost else ""
    )

    return f"""
    <section class="card">
      <h2>{channel['title']} {cost_legend}</h2>
      <p class="card-sub">{channel['subtitle']}</p>
      <table>
        <thead>
          <tr>
            <th>Marca</th>
            <th class="num">{channel['metric']}</th>
            <th class="bar-cell">Distribuição</th>
            {invest_header}
          </tr>
        </thead>
        <tbody>
{chr(10).join(body_rows)}
        </tbody>
      </table>
    </section>
    """.strip()


def fmt_growth(v):
    if v is None:
        return '<span class="muted">N/A</span>'
    sign = "+" if v > 0 else ""
    cls = "pos" if v > 0 else ("neg" if v < 0 else "neu")
    return f'<span class="growth {cls}">{sign}{v:.2f}%</span>'


def write_html(results, output_path):
    from datetime import datetime

    rows_html = []
    for r in results:
        kw = r["keyword"]
        is_focus = kw == FOCUS
        url = URLS.get(kw, "#")
        row_class = ' class="focus"' if is_focus else ""
        rows_html.append(f"""
            <tr{row_class}>
                <td class="brand">
                    <div class="brand-name">{kw.upper()}{' <span class="focus-badge">FOCO</span>' if is_focus else ''}</div>
                    <a class="brand-url" href="{url}" target="_blank" rel="noopener">{url}</a>
                </td>
                <td class="growth-cell">{fmt_growth(r['growth_1y'])}</td>
                <td class="growth-cell">{fmt_growth(r['growth_3y'])}</td>
                <td class="num highlight">{fmt_int(r['2026-03'])}</td>
                <td class="num">{fmt_int(r['2025-03'])}</td>
                <td class="num">{fmt_int(r['2024-03'])}</td>
                <td class="num">{fmt_int(r['2023-03'])}</td>
            </tr>
        """.strip())

    rows_str = "\n".join(rows_html)
    generated_at = datetime.now().strftime("%d/%m/%Y %H:%M")

    channels_html = "\n".join(render_channel(k, v) for k, v in CHANNELS.items())

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Relatório de Marca — Somos Valor</title>
<style>
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0; padding: 40px 20px;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
    color: #e2e8f0;
    min-height: 100vh;
  }}
  .container {{ max-width: 1180px; margin: 0 auto; }}
  header {{ margin-bottom: 32px; }}
  .eyebrow {{
    text-transform: uppercase; letter-spacing: 2px; font-size: 12px;
    color: #94a3b8; margin-bottom: 8px;
  }}
  h1 {{
    font-size: 38px; font-weight: 700; margin: 0 0 8px;
    background: linear-gradient(90deg, #38bdf8, #a78bfa);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    background-clip: text;
  }}
  .subtitle {{ color: #94a3b8; font-size: 15px; max-width: 760px; line-height: 1.5; }}
  .meta {{
    display: flex; gap: 16px; flex-wrap: wrap; margin-top: 16px;
    font-size: 12px; color: #64748b;
  }}
  .meta span strong {{ color: #cbd5e1; }}

  .card {{
    background: rgba(15, 23, 42, 0.6);
    border: 1px solid rgba(148, 163, 184, 0.15);
    border-radius: 16px; padding: 24px;
    margin-bottom: 24px;
    backdrop-filter: blur(10px);
  }}
  .card h2 {{
    margin: 0 0 16px; font-size: 18px; font-weight: 600;
    color: #f1f5f9;
    display: flex; align-items: center; gap: 8px;
  }}
  .card h2::before {{
    content: ""; width: 4px; height: 20px;
    background: linear-gradient(180deg, #38bdf8, #a78bfa);
    border-radius: 2px;
  }}

  table {{
    width: 100%; border-collapse: collapse;
    font-size: 14px;
  }}
  thead th {{
    text-align: left; padding: 14px 12px;
    font-size: 11px; text-transform: uppercase; letter-spacing: 1px;
    color: #94a3b8; font-weight: 600;
    border-bottom: 1px solid rgba(148, 163, 184, 0.2);
    background: rgba(30, 41, 59, 0.5);
  }}
  thead th.num, thead th.growth-cell {{ text-align: right; }}
  tbody td {{
    padding: 18px 12px;
    border-bottom: 1px solid rgba(148, 163, 184, 0.08);
    vertical-align: middle;
  }}
  tbody tr:last-child td {{ border-bottom: none; }}
  tbody tr:hover {{ background: rgba(56, 189, 248, 0.05); }}
  tbody tr.focus {{
    background: linear-gradient(90deg, rgba(167, 139, 250, 0.1), transparent);
    border-left: 3px solid #a78bfa;
  }}

  td.brand {{ min-width: 240px; }}
  .brand-name {{
    font-weight: 700; font-size: 15px; color: #f8fafc;
    margin-bottom: 4px; letter-spacing: 0.5px;
  }}
  .brand-url {{
    font-size: 12px; color: #64748b; text-decoration: none;
  }}
  .brand-url:hover {{ color: #38bdf8; }}
  .focus-badge {{
    display: inline-block; padding: 2px 8px; margin-left: 8px;
    background: linear-gradient(90deg, #a78bfa, #38bdf8);
    color: #0f172a; font-size: 10px; font-weight: 700;
    border-radius: 4px; vertical-align: middle;
  }}

  td.num {{ text-align: right; font-variant-numeric: tabular-nums; font-weight: 500; color: #cbd5e1; }}
  td.num.highlight {{ color: #f8fafc; font-weight: 700; font-size: 16px; }}
  td.growth-cell {{ text-align: right; font-variant-numeric: tabular-nums; }}

  .growth {{
    display: inline-block; padding: 4px 10px;
    border-radius: 6px; font-weight: 600; font-size: 13px;
  }}
  .growth.pos {{ background: rgba(34, 197, 94, 0.15); color: #4ade80; }}
  .growth.neg {{ background: rgba(239, 68, 68, 0.15); color: #f87171; }}
  .growth.neu {{ background: rgba(148, 163, 184, 0.15); color: #cbd5e1; }}
  .muted {{ color: #64748b; font-style: italic; font-size: 12px; }}

  .insights {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 16px; }}
  .insight {{
    background: rgba(30, 41, 59, 0.4);
    border: 1px solid rgba(148, 163, 184, 0.1);
    border-radius: 12px; padding: 18px;
  }}
  .insight .label {{
    text-transform: uppercase; font-size: 10px; letter-spacing: 1.5px;
    color: #64748b; margin-bottom: 6px;
  }}
  .insight .value {{
    font-size: 26px; font-weight: 700; color: #f8fafc; margin-bottom: 4px;
  }}
  .insight .value.pos {{ color: #4ade80; }}
  .insight .value.neg {{ color: #f87171; }}
  .insight .desc {{ font-size: 12px; color: #94a3b8; line-height: 1.5; }}

  .takeaways {{ list-style: none; padding: 0; margin: 0; }}
  .takeaways li {{
    padding: 12px 0 12px 28px; position: relative;
    border-bottom: 1px solid rgba(148, 163, 184, 0.08);
    font-size: 14px; line-height: 1.6; color: #cbd5e1;
  }}
  .takeaways li:last-child {{ border-bottom: none; }}
  .takeaways li::before {{
    content: ""; position: absolute; left: 0; top: 18px;
    width: 8px; height: 8px; border-radius: 50%;
    background: linear-gradient(135deg, #38bdf8, #a78bfa);
  }}
  .takeaways strong {{ color: #f8fafc; }}

  footer {{
    text-align: center; padding: 24px 0;
    color: #64748b; font-size: 12px;
  }}

  .section-title {{
    font-size: 13px; text-transform: uppercase; letter-spacing: 2px;
    color: #94a3b8; margin: 32px 0 16px; font-weight: 600;
  }}
  .card-sub {{
    margin: -8px 0 16px; font-size: 13px; color: #94a3b8;
  }}
  .meta-tag {{
    display: inline-block; padding: 3px 10px; margin-left: 12px;
    background: rgba(56, 189, 248, 0.12);
    color: #38bdf8; font-size: 11px; font-weight: 600;
    border-radius: 4px; vertical-align: middle;
    text-transform: none; letter-spacing: 0.3px;
    border: 1px solid rgba(56, 189, 248, 0.25);
  }}

  .bar-cell {{ width: 35%; min-width: 180px; padding-right: 16px !important; }}
  thead th.bar-cell {{ text-align: left; }}
  .bar-track {{
    width: 100%; height: 8px;
    background: rgba(148, 163, 184, 0.1);
    border-radius: 999px; overflow: hidden;
  }}
  .bar-fill {{
    height: 100%;
    background: linear-gradient(90deg, #38bdf8, #a78bfa);
    border-radius: 999px;
    transition: width 0.4s ease;
  }}
  tr.focus .bar-fill {{
    background: linear-gradient(90deg, #f472b6, #a78bfa);
  }}

  td.invest {{
    color: #fbbf24; font-weight: 600;
    font-variant-numeric: tabular-nums;
  }}
</style>
</head>
<body>
  <div class="container">
    <header>
      <div class="eyebrow">Relatório de Marca · Benchmark Competitivo</div>
      <h1>Somos Valor vs. Mercado</h1>
      <p class="subtitle">Análise de volume de buscas (Google) das marcas concorrentes em janelas de março — 2023 a 2026 — com índices de crescimento de 1 e 3 anos.</p>
      <div class="meta">
        <span><strong>Fonte:</strong> DataForSEO · Google Ads Search Volume</span>
        <span><strong>Localização:</strong> Brasil (BR-2076)</span>
        <span><strong>Gerado em:</strong> {generated_at}</span>
      </div>
    </header>

    <section class="card">
      <h2>Indicadores principais — Somos Valor</h2>
      <div class="insights">
        <div class="insight">
          <div class="label">Volume mar/2026</div>
          <div class="value">{fmt_int(_focus(results)['2026-03'])}</div>
          <div class="desc">buscas mensais pela marca</div>
        </div>
        <div class="insight">
          <div class="label">Crescimento 1 ano</div>
          <div class="value {_growth_cls(_focus(results)['growth_1y'])}">{_growth_text(_focus(results)['growth_1y'])}</div>
          <div class="desc">março/2025 → março/2026</div>
        </div>
        <div class="insight">
          <div class="label">Crescimento 3 anos</div>
          <div class="value {_growth_cls(_focus(results)['growth_3y'])}">{_growth_text(_focus(results)['growth_3y'])}</div>
          <div class="desc">março/2023 → março/2026</div>
        </div>
        <div class="insight">
          <div class="label">Posição no benchmark</div>
          <div class="value">{_rank(results)}/{len(results)}</div>
          <div class="desc">por volume em mar/2026</div>
        </div>
      </div>
    </section>

    <section class="card">
      <h2>Tabela comparativa</h2>
      <table>
        <thead>
          <tr>
            <th>Marca</th>
            <th class="growth-cell">Crescimento 1 ano</th>
            <th class="growth-cell">Crescimento 3 anos</th>
            <th class="num">03/2026</th>
            <th class="num">03/2025</th>
            <th class="num">03/2024</th>
            <th class="num">03/2023</th>
          </tr>
        </thead>
        <tbody>
{rows_str}
        </tbody>
      </table>
    </section>

    <h2 class="section-title">Comparativo por canal</h2>
    {channels_html}

    <footer>
      Relatório gerado automaticamente · Dados Google Ads via DataForSEO
    </footer>
  </div>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)


def _focus(results):
    for r in results:
        if r["keyword"] == FOCUS:
            return r
    return results[0]


def _growth_text(v):
    if v is None:
        return "N/A"
    sign = "+" if v > 0 else ""
    return f"{sign}{v:.1f}%"


def _growth_cls(v):
    if v is None:
        return "neu"
    return "pos" if v > 0 else ("neg" if v < 0 else "neu")


def _rank(results):
    sorted_r = sorted(results, key=lambda r: (r["2026-03"] is None, -(r["2026-03"] or 0)))
    for i, r in enumerate(sorted_r, 1):
        if r["keyword"] == FOCUS:
            return i
    return "—"


if __name__ == "__main__":
    main()
