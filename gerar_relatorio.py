#!/usr/bin/env python3
"""
Voz em Ação — relatório diário da devolutiva do GRUPO (um PDF por cliente).

Fluxo (rodado todo dia pelo GitHub Actions, ver .github/workflows/relatorio-diario.yml):
  1. pergunta ao Apps Script quais clientes tiveram respostas novas  (ação relDados)
  2. para cada cliente com novidade, gera o PDF da roda do grupo       (Chrome sem tela)
  3. devolve o PDF ao Apps Script, que envia o e-mail e guarda no Drive (ação relEnviar)
Se ninguém respondeu desde o último envio, nada é gerado nem enviado.

Só usa a biblioteca padrão do Python. Variáveis de ambiente:
  APPS_SCRIPT_URL, RELATORIO_KEY   (secrets do GitHub)
  FORCAR=1                         (opcional: reenvia mesmo sem novidade)

Teste local, sem rede e sem e-mail:
  python gerar_relatorio.py --xlsx "Registro_Devolutivas.xlsx" --empresa "John Deere" --saida teste.pdf
"""
import argparse, base64, datetime, html, json, math, os, shutil, subprocess, sys, tempfile, time
import urllib.parse, urllib.request
from zoneinfo import ZoneInfo

AQUI = os.path.dirname(os.path.abspath(__file__))
TEMAS = json.load(open(os.path.join(AQUI, "temas.json"), encoding="utf-8"))   # 12 temas, na ordem dos blocos do teste
LOGO = "data:image/png;base64," + base64.b64encode(open(os.path.join(AQUI, "logo.png"), "rb").read()).decode()
BLOCK_CAT = [0, 1, 2, 3, 7, 9, 6, 4, 8, 5, 10, 11]   # bloco de 4 perguntas -> posição na roda
WN = ["Autoconhecimento", "Presença", "Empatia", "Escuta", "Linguagem Corporal", "Comunicação Não Violenta",
      "Assertividade", "Autenticidade", "Conversas Difíceis", "Storytelling", "Feedback", "Confiança"]
W2L = {"Linguagem Corporal": ["LINGUAGEM", "CORPORAL"], "Conversas Difíceis": ["CONVERSAS", "DIFÍCEIS"],
       "Comunicação Não Violenta": ["COMUNICAÇÃO", "NÃO VIOLENTA"]}


# ---------------------------------------------------------------- cálculo
def faixa(a):
    return "prioridade" if a <= 2 else "investir" if a == 3 else "forca"


def calcular(respostas):
    """respostas: lista de pessoas, cada uma com 48 notas (1 a 5)."""
    n = len(respostas)
    avg = [[sum(a[b * 4:b * 4 + 4]) / 4 for b in range(12)] for a in respostas]
    gavg = [sum(p[b] for p in avg) / n for b in range(12)]
    cnt = {"forca": 0, "investir": 0, "prioridade": 0}
    per = [{"forca": 0, "investir": 0, "prioridade": 0} for _ in range(12)]
    for a in respostas:
        for q, v in enumerate(a):
            cnt[faixa(v)] += 1
            per[q // 4][faixa(v)] += 1
    rank = sorted(range(12), key=lambda b: (-gavg[b], b))
    return {"n": n, "gavg": gavg, "glevel": [g * 2 for g in gavg],
            "mn": [min(p[b] for p in avg) for b in range(12)], "mx": [max(p[b] for p in avg) for b in range(12)],
            "pct": {k: round(v / (n * 48) * 100) for k, v in cnt.items()}, "per": per,
            "top": sorted(rank[:3]), "bot": sorted(rank[-3:])}


# ---------------------------------------------------------------- roda
def wheel(strong, size):
    R, band = 300, 46
    pt = lambda r, a: (r * math.cos(math.radians(a)), r * math.sin(math.radians(a)))

    def wedge(r, a0, a1):
        x0, y0 = pt(r, a0)
        x1, y1 = pt(r, a1)
        return f"M0 0L{x0:.1f} {y0:.1f}A{r} {r} 0 0 1 {x1:.1f} {y1:.1f}Z"

    v = R + band + 8
    s = (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="{-v} {-v} {2*v} {2*v}" width="{size}" height="{size}">'
         f'<circle r="{R+band}" fill="#8b7fb8"/><circle r="{R}" fill="#e6e2f1"/>')
    for k in range(12):
        a0 = -120 + 30 * k
        a1 = a0 + 30
        if strong[k]:
            s += f'<path d="{wedge(R*strong[k]/10,a0,a1)}" fill="#5c3e94"/>'
        mid = a0 + 15
        rr = R + band / 2 + 7
        bottom = 0 < mid < 180
        nm = WN[k]
        lines = W2L.get(nm, [nm.upper()])
        fs = 13 if len(lines) > 1 else (15 if len(nm) > 12 else 17)
        ls = fs * 1.1
        for li, ln in enumerate(lines):
            r2 = rr + (-1 if bottom else 1) * (0.5 - li) * ls if len(lines) > 1 else rr
            xa, ya = pt(r2, a1 if bottom else a0)
            xb, yb = pt(r2, a0 if bottom else a1)
            pid = f"p{k}_{li}"
            s += f'<path id="{pid}" d="M{xa:.1f} {ya:.1f}A{r2:.1f} {r2:.1f} 0 0 {0 if bottom else 1} {xb:.1f} {yb:.1f}" fill="none"/>'
            s += (f'<text font-family="Barlow Condensed,Arial Narrow,sans-serif" font-weight="600" font-size="{fs}" letter-spacing=".8" fill="#fff">'
                  f'<textPath href="#{pid}" startOffset="50%" text-anchor="middle">{ln}</textPath></text>')
    for n in range(1, 11):
        s += f'<circle r="{R*n/10}" fill="none" stroke="#8b7fb8" stroke-width="{2.5 if n==10 else 1}"/>'
    for k in range(12):
        x, y = pt(R + band, -120 + 30 * k)
        s += f'<line x1="0" y1="0" x2="{x:.1f}" y2="{y:.1f}" stroke="#8b7fb8" stroke-width="1.5"/>'
    return s + "</svg>"


# ---------------------------------------------------------------- página
def f1(x):
    return f"{x:.1f}".replace(".", ",")


def esc(s):
    return html.escape(str(s))


def d_br(iso):   # 2026-10-29 -> 29/10/2026
    return f"{iso[8:10]}/{iso[5:7]}/{iso[:4]}"


def montar_html(empresa, respostas, primeira, ultima, esperados, hoje):
    c = calcular(respostas)
    N = c["n"]
    parcial = bool(esperados) and N < esperados
    lv = [0.0] * 12
    for b in range(12):
        lv[BLOCK_CAT[b]] = c["glevel"][b]
    roda = wheel(lv, 560)

    linhas = ""
    for b in sorted(range(12), key=lambda b: BLOCK_CAT[b]):
        pt_, n4 = c["per"][b], N * 4
        barra = "".join(f'<i style="width:{pt_[f]/n4*100:.1f}%;background:{cor}"></i>'
                        for f, cor in (("forca", "#0e9ca4"), ("investir", "#b6acd6"), ("prioridade", "#706f6f")))
        linhas += (f'<tr><td><b>{esc(TEMAS[b]["nome"])}</b></td><td class="n">{f1(c["gavg"][b])}</td>'
                   f'<td class="n">{f1(c["glevel"][b])}</td><td class="n">{f1(c["mn"][b])} – {f1(c["mx"][b])}</td>'
                   f'<td><div class="bar">{barra}</div></td></tr>')

    def lista(idx, cor):
        return "".join(f'<div class="it" style="border-left:8px solid {cor}"><b>{esc(TEMAS[b]["nome"])}</b> '
                       f'<span class="sc">média {f1(c["gavg"][b])}</span><br>{esc(TEMAS[b]["habilidade"])}</div>' for b in idx)

    resp_txt = f"{N} de ~{esperados}" if parcial else f"{N}"
    resp_lbl = "Respondentes (resultado parcial)" if parcial else "Respondentes"
    rodape = (f"Resultado parcial com {N} respondentes — será atualizado com as demais respostas." if parcial
              else f"Resultado do grupo com {N} respondentes.")
    rodape2 = f"Resultado parcial com {N} respondentes." if parcial else f"Resultado do grupo com {N} respondentes."
    periodo = d_br(primeira) if primeira == ultima else f"{d_br(primeira)[:5]} a {d_br(ultima)}"
    emp_up = esc(empresa).upper()
    pct = c["pct"]

    return f"""<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>Devolutiva do Grupo</title>
<link href="https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600&family=Barlow+Condensed:wght@600;700&display=swap" rel="stylesheet">
<style>
@page{{size:A4;margin:0}}
*{{box-sizing:border-box}}
body{{margin:0;font-family:Barlow,Arial,sans-serif;color:#3e3434;font-size:11.5pt;line-height:1.35}}
.pg{{width:210mm;height:297mm;padding:16mm 16mm 14mm;page-break-after:always;position:relative;overflow:hidden}}
.pg:last-child{{page-break-after:auto}}
.top{{display:flex;justify-content:space-between;align-items:center;border-bottom:3px solid #5c3e94;padding-bottom:8px;margin-bottom:14px}}
.top img{{height:46px}}
.kick{{font-family:'Barlow Condensed',Arial Narrow,sans-serif;font-weight:700;letter-spacing:.12em;color:#5c3e94;font-size:12pt}}
h1{{font-family:'Barlow Condensed',Arial Narrow,sans-serif;font-weight:700;font-size:30pt;line-height:1;margin:4px 0 6px}}
h2{{font-family:'Barlow Condensed',Arial Narrow,sans-serif;font-weight:700;font-size:17pt;letter-spacing:.05em;margin:16px 0 6px;color:#5c3e94}}
.meta{{display:flex;gap:10px;margin:10px 0}}
.meta div{{flex:1;background:#e6e2f1;border-radius:8px;padding:8px 10px;font-size:10pt}}
.meta b{{display:block;font-size:17pt;color:#5c3e94;font-family:'Barlow Condensed',Arial Narrow,sans-serif}}
.wh{{display:grid;place-items:center;margin:6px 0 4px}}
.leg{{text-align:center;font-size:10pt}}
.leg i{{display:inline-block;width:12px;height:12px;background:#5c3e94;vertical-align:-1px;margin-right:4px}}
table{{width:100%;border-collapse:collapse;font-size:10.5pt}}
th{{text-align:left;font-size:9pt;letter-spacing:.04em;color:#5c3e94;border-bottom:2px solid #5c3e94;padding:4px 6px}}
td{{padding:5px 6px;border-bottom:1px solid #d8d4e6}} td.n{{text-align:center;white-space:nowrap}}
.bar{{display:flex;height:12px;border-radius:6px;overflow:hidden;background:#eee;min-width:120px}} .bar i{{display:block;height:100%}}
.lg2{{font-size:9.5pt;margin-top:6px}} .lg2 i{{display:inline-block;width:11px;height:11px;border-radius:2px;margin:0 3px 0 10px;vertical-align:-1px}}
.it{{background:#f4f2f9;padding:7px 10px;margin:6px 0;font-size:10.5pt}} .sc{{color:#5c3e94;font-size:9.5pt}}
h3{{font-family:'Barlow Condensed',Arial Narrow,sans-serif;font-size:14pt;letter-spacing:.05em;margin:12px 0 2px}}
.note{{font-size:9.5pt;color:#6a6b6e;margin-top:10px}}
.foot{{position:absolute;left:16mm;right:16mm;bottom:8mm;font-size:8.5pt;color:#6a6b6e;display:flex;justify-content:space-between}}
</style></head><body>

<div class="pg">
 <div class="top"><div><div class="kick">PLANNER VOZ EM AÇÃO · TESTE DA COMUNICAÇÃO</div><h1>Devolutiva do Grupo</h1></div><img src="{LOGO}"></div>
 <div class="meta">
  <div>Empresa<b>{esc(empresa)}</b></div>
  <div>{resp_lbl}<b>{resp_txt}</b></div>
  <div>Período das respostas<b>{periodo}</b></div>
 </div>
 <h2>RODA DA COMUNICAÇÃO DO GRUPO</h2>
 <div class="wh">{roda}</div>
 <div class="leg"><i></i>Nível médio do grupo em cada tema (escala 0 a 10 da roda: média das respostas de 1 a 5 × 2)</div>
 <div class="foot"><span>{esc(rodape)}</span><span>Emitido em {hoje} · 1</span></div>
</div>

<div class="pg">
 <div class="top"><div><div class="kick">DEVOLUTIVA DO GRUPO · {emp_up}</div><h1>Resultado por tema</h1></div><img src="{LOGO}"></div>
 <table><tr><th>TEMA</th><th>MÉDIA (1–5)</th><th>NÍVEL NA RODA (0–10)</th><th>MENOR – MAIOR MÉDIA</th><th>DISTRIBUIÇÃO DAS RESPOSTAS</th></tr>{linhas}</table>
 <div class="lg2"><i style="background:#0e9ca4"></i>Forças (4 e 5)<i style="background:#b6acd6"></i>Investir (3)<i style="background:#706f6f"></i>Prioridade (1 e 2)</div>
 <h2>PANORAMA GERAL DAS RESPOSTAS</h2>
 <div class="meta">
  <div>Pontos fortes<b style="color:#0e9ca4">{pct['forca']}%</b>das respostas (4 ou 5)</div>
  <div>Investir<b>{pct['investir']}%</b>das respostas (3)</div>
  <div>Prioridade<b style="color:#706f6f">{pct['prioridade']}%</b>das respostas (1 ou 2)</div>
 </div>
 <p class="note">Média (1–5): média das 4 afirmações de cada tema, depois média entre os respondentes. Nível na roda: média × 2, a mesma conta da devolutiva individual. “Menor – maior média” mostra a variação entre os respondentes. Percentuais calculados sobre {N*48} respostas ({N} pessoas × 48 afirmações).</p>
 <div class="foot"><span>{esc(rodape2)}</span><span>Emitido em {hoje} · 2</span></div>
</div>

<div class="pg">
 <div class="top"><div><div class="kick">DEVOLUTIVA DO GRUPO · {emp_up}</div><h1>Destaques do grupo</h1></div><img src="{LOGO}"></div>
 <h3 style="color:#0e9ca4">PONTOS FORTES DO GRUPO</h3>{lista(c["top"],"#0e9ca4")}
 <h3 style="color:#706f6f;margin-top:18px">ONDE INVESTIR E DESENVOLVER</h3>{lista(c["bot"],"#706f6f")}
 <p class="note">Os três temas com maior média e os três com menor média do grupo, no mesmo critério da devolutiva individual. Este relatório traz apenas resultados agrupados, sem identificação das pessoas.</p>
 <div class="foot"><span>{esc(rodape2)}</span><span>Emitido em {hoje} · 3</span></div>
</div>
</body></html>"""


# ---------------------------------------------------------------- PDF
def achar_chrome():
    if os.environ.get("CHROME"):
        return os.environ["CHROME"]
    for nome in ("google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "msedge"):
        p = shutil.which(nome)
        if p:
            return p
    for p in (r"C:\Program Files\Google\Chrome\Application\chrome.exe",
              r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
              r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
              r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"):
        if os.path.exists(p):
            return p
    sys.exit("Chrome/Edge não encontrado (defina a variável CHROME).")


def html_para_pdf(doc, saida):
    with tempfile.TemporaryDirectory() as tmp:
        arq = os.path.join(tmp, "r.html")
        open(arq, "w", encoding="utf-8").write(doc)
        if os.path.exists(saida):
            os.remove(saida)
        cmd = [achar_chrome(), "--headless=new", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
               f"--user-data-dir={os.path.join(tmp, 'perfil')}", "--virtual-time-budget=20000",
               f"--print-to-pdf={saida}", "file:///" + arq.replace("\\", "/")]
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
        ultimo, t0 = -1, time.time()
        while time.time() - t0 < 60:   # o Chrome às vezes termina o processo antes de fechar o arquivo
            tam = os.path.getsize(saida) if os.path.exists(saida) else -1
            if tam > 0 and tam == ultimo:
                break
            ultimo = tam
            time.sleep(1.5)
    if not os.path.exists(saida) or os.path.getsize(saida) < 5000:
        sys.exit("Falha ao gerar o PDF.")
    with open(saida, "rb") as f:
        if f.read(4) != b"%PDF":
            sys.exit("Arquivo gerado não é um PDF.")


# ---------------------------------------------------------------- Apps Script
def chamar(url, params):
    # POST: a chave não aparece no endereço. O Apps Script responde 302; o urllib repete como GET no endereço de resposta.
    req = urllib.request.Request(url, data=urllib.parse.urlencode(params).encode())
    with urllib.request.urlopen(req, timeout=180) as r:
        txt = r.read().decode("utf-8", "replace")
    try:
        return json.loads(txt)
    except ValueError:
        sys.exit("Resposta inesperada do Apps Script (confira a URL e se a implantação foi atualizada): " + txt[:200])


def hoje_br():
    return datetime.datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y")


def modo_nuvem():
    url, chave = os.environ.get("APPS_SCRIPT_URL", "").strip(), os.environ.get("RELATORIO_KEY", "").strip()
    if not url or not chave:
        sys.exit("Faltam os secrets APPS_SCRIPT_URL e/ou RELATORIO_KEY.")
    dados = chamar(url, {"action": "relDados", "key": chave, "todos": "1" if os.environ.get("FORCAR") == "1" else "0"})
    if not dados.get("ok"):
        sys.exit("relDados falhou: " + str(dados.get("error")))
    clientes = dados.get("clientes", [])
    novos = [c for c in clientes if c.get("mudou")]
    print(f"{len(clientes)} cliente(s) com respostas; {len(novos)} com novidade.")
    falhou = False
    for c in novos:
        nome = c["empresa"]
        pdf = os.path.join(tempfile.gettempdir(), "grupo_" + str(abs(hash(nome))) + ".pdf")
        doc = montar_html(nome, c["respostas"], c["primeira"], c["ultima"], c.get("esperados"), hoje_br())
        html_para_pdf(doc, pdf)
        b64 = base64.b64encode(open(pdf, "rb").read()).decode()
        r = chamar(url, {"action": "relEnviar", "key": chave, "empresa": nome, "sig": c["sig"], "n": c["n"],
                         "data": hoje_br(), "pdf": b64})
        print(f"- {nome}: {'enviado' if r.get('ok') else 'ERRO ' + str(r.get('error'))}")
        falhou = falhou or not r.get("ok")
    if falhou:
        sys.exit(1)
    if not novos:
        print("Sem novidades: nenhum e-mail enviado.")


def _compacto(s):
    import re, unicodedata
    t = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode().lower()
    t = re.sub(r"\b(ltda|sa|eireli|inc|corp|cia|brasil|brazil|do)\b", " ", t).replace("s/a", " ")
    return re.sub(r"[^a-z0-9]", "", t)


def modo_teste(xlsx, empresa, saida):
    import openpyxl
    ws = openpyxl.load_workbook(xlsx, data_only=True).active
    h = [c.value for c in ws[1]]
    rows = [dict(zip(h, [c.value for c in r])) for r in ws.iter_rows(min_row=2) if r[1].value]
    alvo = _compacto(empresa)
    # mesma ideia do Apps Script: nome da empresa com variações (Brasil, Ltda, acento, maiúscula…) ou domínio do e-mail
    rows = [r for r in rows if (alvo in _compacto(r.get("Empresa")) or alvo in str(r.get("Email") or "").lower().split("@")[-1])]
    ultimos = {}
    for r in rows:   # quem refez o teste vale pela resposta mais recente
        ultimos[str(r["Email"]).strip().lower()] = r
    rows = list(ultimos.values())
    resp = [[int(r[f"Q{q}"]) for q in range(1, 49)] for r in rows]
    datas = sorted(str(r["Data_Hora"])[:10] for r in rows)
    doc = montar_html(empresa, resp, datas[0], datas[-1], 22, hoje_br())
    html_para_pdf(doc, saida)
    print(f"PDF de teste ({len(resp)} respondentes): {saida}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx")
    ap.add_argument("--empresa", default="John Deere")
    ap.add_argument("--saida", default="relatorio_teste.pdf")
    a = ap.parse_args()
    modo_teste(a.xlsx, a.empresa, os.path.abspath(a.saida)) if a.xlsx else modo_nuvem()
