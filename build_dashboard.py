"""Gera dashboard_redes.html: dashboard autocontido da UEA nas 3 redes CAPES Global.

Parte do painel de setembro/2027 (painel_redes.html, em Relatorios/REUNIAO PATRICIA,SERGIO,SCOOT)
e o enriquece com o mapeamento completo dos Anexos I das propostas: todas as áreas
temáticas, todas as IES e todos os PPGs de cada rede, com nota CAPES.

Uso: python3 build_dashboard.py
Fontes:
  - CAPES_Global_Mapeamento_3_Redes_PPGs_IES.xlsm (abas PANORAMA e <REDE> - PPGs);
  - valores: fontes/Balanco_CAPES_Global_2027_Orcamentos_Ajustados.xlsm - Google Planilhas.pdf
    (aba Setembro 2027 | Bolsas a atender), lido com pdftotext;
  - comitê UEA (temáticas, coordenadores, PPGs cadastrados): estrutura_capes_global.json
    do Gerenciador PROPESP.
"""
import base64
import html
import json
import math
import re
import subprocess
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

AQUI = Path(__file__).resolve().parent
VENDOR = AQUI / 'vendor'
XLSM = AQUI / 'CAPES_Global_Mapeamento_3_Redes_PPGs_IES.xlsm'
PDF = AQUI / 'fontes' / 'Balanco_CAPES_Global_2027_Orcamentos_Ajustados.xlsm - Google Planilhas.pdf'
DJ = Path('/Users/macbookpro/Documents/POSDOC - MAC/UEA/GERENCIADOR PROPESP/SISTEMA NOVO - DJANGO')

REDES = ['UNESP', 'UFMG', 'UNICAMP']
NOME_REDE = {'UNESP': 'Rede Conecta Mundi (RCM)', 'UFMG': 'Rede INTER-NÓS', 'UNICAMP': 'SustentNET'}
ABA = {r: f'{r} - PPGs' for r in REDES}
COR = {'UNESP': '#2a78d6', 'UFMG': '#eb6834', 'UNICAMP': '#1baf7a'}
CLARO = {'UNESP': '#e6f0fc', 'UFMG': '#fdeee6', 'UNICAMP': '#e3f6ee'}
LOGO_ARQ = {'UNESP': ('unesp.svg', 'image/svg+xml'), 'UFMG': ('ufmg.png', 'image/png'),
            'UNICAMP': ('unicamp.webp', 'image/webp'), 'UEA': ('uea.png', 'image/png')}

# Sede/reitoria de cada universidade: (cidade, UF, lat, lon, lado do rótulo)
LOCAL = {
    'UEA': ('Manaus', 'AM', -3.10, -60.02, 'd'),
    'UNESP': ('São Paulo', 'SP', -23.55, -46.64, 'd'),
    'UFS': ('São Cristóvão', 'SE', -10.93, -37.10, 'e'),
    'UEMS': ('Dourados', 'MS', -22.22, -54.81, 'e'),
    'UNEB': ('Salvador', 'BA', -12.97, -38.50, 'd'),
    'UNIOESTE': ('Cascavel', 'PR', -24.96, -53.46, 'e'),
    'UFMG': ('Belo Horizonte', 'MG', -19.87, -43.96, 'd'),
    'UFCSPA': ('Porto Alegre', 'RS', -30.03, -51.22, 'e'),
    'UFRN': ('Natal', 'RN', -5.84, -35.20, 'e'),
    'UNIFAL-MG': ('Alfenas', 'MG', -21.43, -45.95, 'e'),
    'UNICAMP': ('Campinas', 'SP', -22.82, -47.07, 'd'),
    'UDESC': ('Florianópolis', 'SC', -27.59, -48.55, 'd'),
    'UNEMAT': ('Cáceres', 'MT', -16.07, -57.68, 'e'),
    'UEFS': ('Feira de Santana', 'BA', -12.27, -38.97, 'd'),
}
SIGLA_PDF = {'CLIAMB': 'PPG CLIAMB'}

# Temática do comitê UEA (estrutura) -> área temática do Anexo I
TEMA_AREA = {
    'UNESP-T1': 'Educação, linguagem, arte e cultura para a construção de uma sociedade plural',
    'UNESP-T2': 'Tecnologias convergentes, materiais avançados e transformação digital',
    'UNESP-T3': 'Saúde Única: interconexão humana, animal, vegetal e ambiental',
    'UNESP-T4': 'Desenvolvimento sustentável: biodiversidade, transição energética e mudanças climáticas',
    'UFMG-T1': 'Justiça, Governança e Instituições Eficazes',
    'UFMG-T2': 'Transformação Digital e Novas Tecnologias',
    'UFMG-T3': 'Saúde e Bem-estar',
    'UFMG-T4': 'Sustentabilidade e Transição Energética',
    'UNICAMP-T1': 'Saúde, Natureza e Sociedade',
    'UNICAMP-T2': 'Culturas, Saberes e Educação',
    'UNICAMP-T3': 'Biodiversidade, Energia e Transformação Territorial Sustentável',
    'UNICAMP-T4': 'Direitos Humanos, Justiça Social e Equidade',
}
# PPG da UEA no Anexo I (nome normalizado) -> sigla usada no comitê e no balanço
SIGLA_UEA = {
    'LETRAS E ARTES': 'PPGLA', 'EDUCACAO': 'PPGED', 'EDUCACAO NA AMAZONIA': 'EDUCANORTE',
    'SEGURANCA PUBLICA, CIDADANIA E DIREITOS HUMANOS': 'PPGSP', 'MEDICINA TROPICAL': 'PPGMT',
    'SAUDE DA FAMILIA': 'PROFSAÚDE', 'HEMATOLOGIA': 'PPGH',
    'GESTAO E REGULACAO DE RECURSOS HIDRICOS - PROFAGUA': 'PROFÁGUA',
    'MULTICENTRICO EM BIOQUIMICA E BIOLOGIA MOLECULAR': 'PPG PMBqBM',
    'BIODIVERSIDADE E BIOTECNOLOGIA - REDE BIONORTE': 'BIONORTE',
    'CLIMA E AMBIENTE - INPA - UEA': 'PPG CLIAMB',
    'BIOTECNOLOGIA E RECURSOS NATURAIS DA AMAZONIA': 'PPGMBT',
    'EDUCACAO EM CIENCIAS E MATEMATICA - UFMT - UFPA - UEA': 'REAMEC',
    'CIENCIAS HUMANAS': 'PPGICH', 'EDUCACAO EM CIENCIAS NA AMAZONIA': 'PPGEEC',
    'MATEMATICA EM REDE NACIONAL': 'PROFMAT', 'DIREITO AMBIENTAL': 'PPGDA',
    'ENFERMAGEM EM SAUDE PUBLICA': 'PPG PROENSP', 'CIENCIAS APLICADAS A DERMATOLOGIA': 'PPGCAD',
    'SAUDE PUBLICA NA AMAZONIA': 'PPG DASPAM',
}

e = html.escape


def sem_acento(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').upper().strip()


def capitaliza(s):
    """Nomes de PPG vêm ora em CAIXA ALTA, ora não; padroniza para exibição."""
    if s != s.upper():
        return s
    minus = {'e', 'de', 'da', 'do', 'das', 'dos', 'em', 'na', 'no', 'para', 'a', 'à', 'ao', 'com'}
    pal = s.lower().split(' ')
    return ' '.join(p if (i and p in minus) else (p.upper() if p in ('ufmt', 'ufpa', 'uea', 'inpa', 'ii') else p[:1].upper() + p[1:])
                    for i, p in enumerate(pal))


def logo(sigla):
    arq, mime = LOGO_ARQ[sigla]
    dados = (DJ / 'static/img/instituicoes' / arq).read_bytes()
    return f'data:{mime};base64,{base64.b64encode(dados).decode()}'


def br(v, casas=2):
    s = f'{v:,.{casas}f}'
    return s.replace(',', 'X').replace('.', ',').replace('X', '.')


def mil(v):
    if v >= 1_000_000:
        return f'R$ {br(v / 1_000_000, 2)} mi'
    return f'R$ {br(v / 1000, 1)} mil'


def num(s):
    return float(s.replace('R$', '').replace('.', '').replace(',', '.').strip())


# ---------------------------------------------------------------- planilha (Anexos I)
wb = openpyxl.load_workbook(XLSM, data_only=True)
pan = list(wb['PANORAMA'].iter_rows(values_only=True))

SIGLA_IES, PAPEL, IES_REDE = {}, {}, defaultdict(list)
i0 = next(i for i, r in enumerate(pan) if r[0] and str(r[0]).startswith('UNIVERSIDADES VINCULADAS'))
for r in pan[i0 + 2:]:
    if r[0] and str(r[0]).startswith('RELAÇÃO COMPLETA'):
        break
    if not r[0]:
        continue
    rede = next(k for k, v in NOME_REDE.items() if v == r[0])
    SIGLA_IES[r[1]] = r[2]
    PAPEL[rede, r[2]] = r[3]
    IES_REDE[rede].append(r[2])
for r in REDES:  # coordenadora, UEA, demais na ordem da planilha
    IES_REDE[r] = [r, 'UEA'] + [u for u in IES_REDE[r] if u not in (r, 'UEA')]

vinc = []  # todas as vinculações (rede, área, IES, PPG)
AREAS = defaultdict(list)
for r in REDES:
    for row in wb[ABA[r]].iter_rows(min_row=5, values_only=True):
        if not row[0]:
            continue
        area, inst, ppg, cod, nota, pag = row
        sig = SIGLA_IES[inst]
        if area not in AREAS[r]:
            AREAS[r].append(area)
        sigla_uea = SIGLA_UEA[sem_acento(ppg)] if sig == 'UEA' else None
        vinc.append({'rede': r, 'area': area, 'ies': sig, 'ppg': capitaliza(ppg), 'cod': cod,
                     'nota': int(nota), 'pag': int(pag), 'sigla': sigla_uea})
for r in REDES:
    assert set(TEMA_AREA[t] for t in TEMA_AREA if t.startswith(r + '-')) <= set(AREAS[r]), r

# conferência com o PANORAMA
for r in REDES:
    vr = [v for v in vinc if v['rede'] == r]
    lin = next(x for x in pan if x[0] == NOME_REDE[r] and isinstance(x[3], int) and isinstance(x[4], int) and isinstance(x[5], int) and x[2] and ',' in x[2] and x[1] and 'Universidade' in x[1])
    assert (len(vr), len({v['cod'] for v in vr})) == (lin[4], lin[5]), (r, len(vr), lin)

# ---------------------------------------------------------------- comitê UEA
estr = json.load(open(DJ / 'apps/estrutura/dados/estrutura_capes_global.json', encoding='utf-8'))
TEMAS = {t['codigo']: t for t in estr['tematicas']}
AREA_TEMA = {(c.split('-')[0], a): c for c, a in TEMA_AREA.items()}
cad = defaultdict(list)  # código do tema -> PPGs cadastrados
for p in estr['ppgs']:
    cad[p['temaId']].append(p)
cad_sigla = {p['sigla']: p for p in estr['ppgs']}

# ---------------------------------------------------------------- valores (PDF)
texto = subprocess.run(['pdftotext', '-layout', str(PDF), '-'], capture_output=True, text=True, check=True).stdout
pedidos, anterior = [], ''
for ln in texto.splitlines():
    m = re.match(r'^(UNICAMP|UNESP|UFMG)\s{2,}(.*)$', ln)
    if not m:
        anterior = ln.strip() if ln[:1].isspace() and ln.strip() else ''
        continue
    cols = re.split(r'\s{2,}', m.group(2).strip())
    tema, ppg, modal, destino, meses, qtde, unit, resto = cols[0], cols[1], cols[2], cols[3], cols[4], cols[5], cols[6], ' '.join(cols[7:])
    if anterior:
        tema = f'{anterior} {tema}'
    anterior = ''
    mt = re.match(r'(R\$\s*[\d.]+,\d{2})\s*(.*)$', resto)

    def chave(s):
        return set(re.findall(r'\w{4,}', s.lower()))
    cand = [t for t in TEMAS.values() if t['rede'] == m.group(1)]
    tcod = max(cand, key=lambda t: len(chave(t['nome']) & chave(tema)))['codigo']
    pedidos.append({'rede': m.group(1), 'tema': tcod, 'ppg': SIGLA_PDF.get(ppg, ppg), 'modalidade': modal,
                    'destino': destino, 'qtde': int(qtde), 'unit': num(unit), 'total': num(mt.group(1))})
for p in pedidos:
    assert abs(p['qtde'] * p['unit'] - p['total']) < 0.01, p
assert len(pedidos) == 24, len(pedidos)

val_rede, bol_rede = defaultdict(float), defaultdict(int)
val_tema = defaultdict(float)
val_ppg, bol_ppg = defaultdict(float), defaultdict(int)  # (tema, sigla)
modal_rede = defaultdict(lambda: defaultdict(float))
for p in pedidos:
    val_rede[p['rede']] += p['total']
    bol_rede[p['rede']] += p['qtde']
    val_tema[p['tema']] += p['total']
    val_ppg[p['tema'], p['ppg']] += p['total']
    bol_ppg[p['tema'], p['ppg']] += p['qtde']
    modal_rede[p['rede']][p['modalidade']] += p['total']
total_geral = sum(val_rede.values())

# ---------------------------------------------------------------- indicadores
def distintos(vs):
    return {v['cod']: v for v in vs}


por_rede = {r: [v for v in vinc if v['rede'] == r] for r in REDES}
cel = defaultdict(list)  # (rede, área, IES) -> vinculações
for v in vinc:
    cel[v['rede'], v['area'], v['ies']].append(v)


def ies_stats(r, u):
    vs = [v for v in por_rede[r] if v['ies'] == u]
    d = distintos(vs)
    notas = [v['nota'] for v in d.values()]
    return {'n': len(d), 'vinc': len(vs), 'areas': len({v['area'] for v in vs}),
            'media': sum(notas) / len(notas) if notas else 0,
            'top': sum(1 for n in notas if n >= 6), 'notas': Counter(notas)}


STATS = {(r, u): ies_stats(r, u) for r in REDES for u in IES_REDE[r]}

# ---------------------------------------------------------------- conferência comitê × anexo
def conferencia(r):
    itens = []
    anexo_uea = defaultdict(set)  # sigla -> áreas do anexo
    nota_anexo = {}
    for v in por_rede[r]:
        if v['ies'] == 'UEA':
            anexo_uea[v['sigla']].add(v['area'])
            nota_anexo[v['sigla']] = v['nota']
    temas_r = [c for c in TEMAS if c.startswith(r + '-')]
    for c in sorted(temas_r):
        area = TEMA_AREA[c]
        for p in cad[c]:
            if p['sigla'] not in anexo_uea:
                itens.append(('alerta', f'<b>{e(p["sigla"])}</b> está no comitê T{c[-1]}, mas não consta como PPG da UEA no Anexo I desta rede.'))
            elif area not in anexo_uea[p['sigla']]:
                itens.append(('alerta', f'<b>{e(p["sigla"])}</b> está no comitê T{c[-1]} ({e(area)}), mas o Anexo I o lista em: {e("; ".join(sorted(anexo_uea[p["sigla"]])))}.'))
            if p['sigla'] in nota_anexo and p['notaCapes'] != nota_anexo[p['sigla']]:
                itens.append(('nota', f'<b>{e(p["sigla"])}</b>: nota {p["notaCapes"]} no cadastro do comitê, {nota_anexo[p["sigla"]]} no Anexo I.'))
    for (t, sig), v in val_ppg.items():
        if t.startswith(r + '-') and sig in anexo_uea and TEMA_AREA[t] not in anexo_uea[sig]:
            itens.append(('alerta', f'Pedido de <b>{mil(v)}</b> de <b>{e(sig)}</b> lançado na T{t[-1]}, área em que o Anexo I não lista esse PPG.'))
        elif t.startswith(r + '-') and sig in anexo_uea and not any(p['sigla'] == sig for p in cad[t]):
            itens.append(('info', f'<b>{e(sig)}</b> tem pedido de {mil(v)} na T{t[-1]} e está cadastrado em outra temática do comitê — o Anexo I confirma o PPG nesta área.'))
    for a in AREAS[r]:
        tem_uea = bool(cel[r, a, 'UEA'])
        if (r, a) not in AREA_TEMA:
            itens.append(('info', f'Área <b>{e(a)}</b>: {"UEA listada no anexo, mas" if tem_uea else "UEA não participa e"} não há temática do comitê UEA.'))
        elif not tem_uea:
            itens.append(('alerta', f'Área <b>{e(a)}</b> tem temática no comitê (T{AREA_TEMA[r, a][-1]}), mas o Anexo I não lista PPG da UEA nela.'))
    ordem = {'alerta': 0, 'nota': 1, 'info': 2}
    return sorted(itens, key=lambda x: ordem[x[0]])


# ---------------------------------------------------------------- mapa
MAPA_W, MAPA_H = 560, 540
LON0, LON1, LAT0, LAT1 = -74.2, -34.6, -33.9, 5.4
KX = math.cos(math.radians(-14))
esc = min(MAPA_W / ((LON1 - LON0) * KX), MAPA_H / (LAT1 - LAT0))
OX = (MAPA_W - (LON1 - LON0) * KX * esc) / 2
OY = (MAPA_H - (LAT1 - LAT0) * esc) / 2


def proj(lon, lat):
    return OX + (lon - LON0) * KX * esc, OY + (LAT1 - lat) * esc


def caminho_estados():
    geo = json.load(open(VENDOR / 'br.geojson', encoding='utf-8'))
    out = []
    for f in geo['features']:
        g = f['geometry']
        polys = g['coordinates'] if g['type'] == 'MultiPolygon' else [g['coordinates']]
        d = []
        for poly in polys:
            for anel in poly[:1]:
                pts, ult = [], None
                for lon, lat in anel:
                    x, y = proj(lon, lat)
                    if ult is None or abs(x - ult[0]) + abs(y - ult[1]) > 1.2:
                        pts.append((x, y))
                        ult = (x, y)
                if len(pts) > 3:
                    d.append('M' + 'L'.join(f'{x:.1f},{y:.1f}' for x, y in pts) + 'Z')
        out.append((f['properties']['sigla'], ''.join(d)))
    return out


ESTADOS = caminho_estados()


def estrela(x, y, s):
    return ' '.join(f'{x + s * math.cos(math.radians(-90 + i * 36)) * (1 if i % 2 == 0 else .45):.1f},'
                    f'{y + s * math.sin(math.radians(-90 + i * 36)) * (1 if i % 2 == 0 else .45):.1f}' for i in range(10))


def mapa(r):
    univs = IES_REDE[r]
    ufs = {LOCAL[u][1] for u in univs}
    estados = ''.join(f'<path d="{d}" class="uf{" uf-on" if uf in ufs else ""}"/>' for uf, d in ESTADOS)
    ux, uy = proj(LOCAL['UEA'][3], LOCAL['UEA'][2])
    nmax = max(STATS[r, u]['n'] for u in univs)
    arcos, pontos = [], []
    for u in univs:
        cid, uf, lat, lon, lado = LOCAL[u]
        x, y = proj(lon, lat)
        st = STATS[r, u]
        if u != 'UEA':
            mx, my = (ux + x) / 2, (uy + y) / 2
            dx, dy = x - ux, y - uy
            arcos.append(f'<path d="M{ux:.1f},{uy:.1f} Q{mx - dy * .18:.1f},{my + dx * .18:.1f} {x:.1f},{y:.1f}" class="arco"/>')
        anc, tx = ('start', x + 15) if lado == 'd' else ('end', x - 15)
        raio = 6 + 12 * math.sqrt(st['n'] / nmax)
        tip = f'data-tip="{e(u)} · {e(cid)}/{uf}|{st["n"]} PPGs distintos em {st["areas"]} áreas"'
        if u == 'UEA':
            marca = (f'<circle cx="{x:.1f}" cy="{y:.1f}" r="{raio + 7:.1f}" class="pulso"/>'
                     f'<circle cx="{x:.1f}" cy="{y:.1f}" r="{raio:.1f}" fill="#0b6b3a" stroke="#fff" stroke-width="2" {tip}/>')
            cls = 'rot rot-uea'
        elif u == r:
            marca = (f'<circle cx="{x:.1f}" cy="{y:.1f}" r="{raio:.1f}" fill="{COR[r]}" fill-opacity=".25" stroke="{COR[r]}" stroke-width="2" {tip}/>'
                     f'<polygon points="{estrela(x, y, 10)}" fill="{COR[r]}" stroke="#fff" stroke-width="2" pointer-events="none"/>')
            cls = 'rot rot-lider'
        else:
            marca = f'<circle cx="{x:.1f}" cy="{y:.1f}" r="{raio:.1f}" fill="{COR[r]}" fill-opacity=".8" stroke="#fff" stroke-width="2" {tip}/>'
            cls = 'rot'
        pontos.append(f'{marca}<text x="{tx:.1f}" y="{y + 5:.1f}" text-anchor="{anc}" class="{cls}">{e(u)}<tspan class="rot-n" dx="5">{st["n"]}</tspan></text>')
    return (f'<svg class="mapa" viewBox="0 0 {MAPA_W} {MAPA_H}" role="img" aria-label="Mapa das universidades da rede {e(NOME_REDE[r])}">'
            f'<g>{estados}</g><g>{"".join(arcos)}</g><g>{"".join(pontos)}</g></svg>')


# ---------------------------------------------------------------- blocos
NOTA_COR = {3: '#dce3ee', 4: '#a8b8d2', 5: '#6f88b4', 6: '#3d5a8e', 7: '#14213d'}


def barra_notas(cnt, rotulo=True):
    tot = sum(cnt.values())
    segs = ''.join(
        f'<span style="flex:{cnt[n]};background:{NOTA_COR[n]};color:{"#fff" if n >= 5 else "var(--ink)"}" '
        f'data-tip="Nota {n}|{cnt[n]} PPGs · {br(cnt[n] / tot * 100, 0)}%">{cnt[n] if rotulo and cnt[n] / tot > .07 else ""}</span>'
        for n in range(3, 8) if cnt[n])
    return f'<div class="notas">{segs}</div>'


LEG_NOTAS = ('<div class="leg-notas">Nota CAPES ' + ''.join(
    f'<span><i style="background:{NOTA_COR[n]}"></i>{n}</span>' for n in range(3, 8)) + '</div>')


def tabela_ies(r):
    nmax = max(STATS[r, u]['n'] for u in IES_REDE[r])
    linhas = []
    for u in IES_REDE[r]:
        cid, uf = LOCAL[u][:2]
        st = STATS[r, u]
        papel = 'coordenadora' if u == r else ('associada' if u != 'UEA' else 'associada · nós')
        ico = ('<i class="dot dot-uea"></i>' if u == 'UEA' else
               f'<i class="dot-lider" style="color:{COR[r]}">★</i>' if u == r else f'<i class="dot" style="background:{COR[r]}"></i>')
        val = f'<b class="uv">{mil(val_rede[r])}</b>' if u == 'UEA' else '<span class="uv mudo">—</span>'
        linhas.append(f'''<tr class="{"tr-uea" if u == "UEA" else ""}">
          <td>{ico}</td><td><b class="us">{e(u)}</b><span class="uc">{e(cid)} · {uf} · {papel}</span></td>
          <td class="bar-c"><div class="hbar"><span style="width:{st["n"] / nmax * 100:.1f}%"></span></div><b>{st["n"]}</b></td>
          <td class="c">{st["areas"]}/{len(AREAS[r])}</td>
          <td class="c">{br(st["media"], 1)}</td>
          <td class="notas-c">{barra_notas(st["notas"], False)}</td>
          <td class="r">{val}</td></tr>''')
    return f'''<table class="t-ies"><thead><tr><th></th><th>IES</th><th>PPGs distintos</th><th class="c">Áreas</th>
      <th class="c">Nota média</th><th>Perfil de notas</th><th class="r">Previsto UEA</th></tr></thead><tbody>{"".join(linhas)}</tbody></table>'''


def matriz(r):
    univs = IES_REDE[r]
    vmax = max(len(distintos(cel[r, a, u])) for a in AREAS[r] for u in univs)
    cab = ''.join(f'<th class="{"th-uea" if u == "UEA" else ""}">{e(u)}</th>' for u in univs)
    linhas = []
    for i, a in enumerate(AREAS[r]):
        tcod = AREA_TEMA.get((r, a))
        badge = (f'<span class="tb" data-tip="Comitê UEA T{tcod[-1]}|{e(TEMAS[tcod]["nome"])}">T{tcod[-1]}</span>' if tcod
                 else '<span class="tb tb-0" data-tip="Sem temática no comitê UEA">—</span>')
        cels = []
        for u in univs:
            d = distintos(cel[r, a, u])
            n = len(d)
            if not n:
                cels.append(f'<td class="hm hm-0{" td-uea" if u == "UEA" else ""}">·</td>')
                continue
            k = .14 + .86 * n / vmax
            lista = '|'.join(f'{v["ppg"]} ({v["nota"]})' for v in sorted(d.values(), key=lambda v: (-v['nota'], v['ppg'])))
            cor = '#0b6b3a' if u == 'UEA' else COR[r]
            cels.append(f'<td class="hm{" td-uea" if u == "UEA" else ""}" style="--k:{k * 100:.0f}%;--hc:{cor};color:{"#fff" if k > .5 else "var(--ink)"}" '
                        f'data-tip="{e(u)} · {e(a)}|{e(lista)}" data-filtro="{e(r)}§{e(a)}§{e(u)}">{n}</td>')
        tot = len(distintos([v for u in univs for v in cel[r, a, u]]))
        val = val_tema.get(tcod, 0) if tcod else 0
        linhas.append(f'<tr><th class="area">{badge}<span>{e(a)}</span></th>{"".join(cels)}'
                      f'<td class="tot">{tot}</td><td class="val">{mil(val) if val else "—"}</td></tr>')
    return (f'<div class="rolar"><table class="matriz"><thead><tr><th class="area">Área temática do Anexo I</th>{cab}'
            f'<th>Total</th><th>Previsto UEA</th></tr></thead><tbody>{"".join(linhas)}</tbody></table></div>')


def card_area(r, a, idx):
    tcod = AREA_TEMA.get((r, a))
    d_uea = distintos(cel[r, a, 'UEA'])
    v = val_tema.get(tcod, 0) if tcod else 0
    pct = v / val_rede[r] * 100 if v else 0
    if tcod:
        t = TEMAS[tcod]
        coord = ' · '.join(re.sub(r'^Prof[a]?\. Dr[a]?\. ', '', m) for m in t['membros'])
        sub = f'<span class="tb">T{tcod[-1]}</span> Comitê UEA · Coord. {e(coord)}'
        valor = (f'<b>{mil(v)}</b><span>{br(pct, 0)}% da rede</span>' if v
                 else f'<b class="ad">{"a definir" if t["statusOrcamento"] == "A_DEFINIR" else "sem pedido"}</b><span>set/2027</span>')
    else:
        sub = '<span class="tb tb-0">—</span> Sem temática no comitê UEA'
        valor = '<b class="ad">—</b>'
    # chips da UEA: PPGs do anexo nesta área + pedidos lançados aqui
    siglas = {x['sigla']: x for x in d_uea.values()}
    if tcod:
        for (t2, sig), _ in val_ppg.items():
            if t2 == tcod and sig not in siglas:
                siglas[sig] = None
    chips = []
    for sig, x in sorted(siglas.items(), key=lambda kv: -val_ppg.get((tcod, kv[0]), 0)):
        pv, nb = val_ppg.get((tcod, sig), 0), bol_ppg.get((tcod, sig), 0)
        nota = f'<s>{x["nota"]}</s>' if x else '<s class="fora" title="Não listado nesta área no Anexo I">!</s>'
        tip = f'data-tip="{e(sig)}|{e(x["ppg"]) if x else "Não listado nesta área no Anexo I"}{" · nota " + str(x["nota"]) if x else ""}"'
        if pv:
            chips.append(f'<span class="chip" {tip}>{nota}<b>{e(sig)}</b><i>{mil(pv)}</i><u>{nb}×</u></span>')
        else:
            chips.append(f'<span class="chip chip-0" {tip}>{nota}<b>{e(sig)}</b></span>')
    uea_html = ''.join(chips) or '<span class="mudo">UEA não participa desta área no Anexo I</span>'
    outras = []
    for u in IES_REDE[r]:
        if u == 'UEA':
            continue
        d = distintos(cel[r, a, u])
        if not d:
            outras.append(f'<details class="ies ies-0"><summary><b>{e(u)}</b><span>—</span></summary></details>')
            continue
        itens = ''.join(f'<li><s>{x["nota"]}</s>{e(x["ppg"])}</li>' for x in sorted(d.values(), key=lambda x: (-x['nota'], x['ppg'])))
        outras.append(f'<details class="ies"><summary><b>{e(u)}</b><span>{len(d)}</span></summary><ul>{itens}</ul></details>')
    n_tot = len(distintos([x for u in IES_REDE[r] for x in cel[r, a, u]]))
    n_ies = sum(1 for u in IES_REDE[r] if cel[r, a, u])
    return f'''
    <article class="area-card{" sem-uea" if not d_uea else ""}">
      <div class="tema-h">
        <span class="num">A{idx}</span>
        <div class="tema-t"><h4>{e(a)}</h4><span class="coord">{sub}</span></div>
        <div class="tema-v">{valor}</div>
      </div>
      <div class="barra-tema"><span style="width:{pct:.1f}%"></span></div>
      <div class="linha-u"><span class="lbl lbl-uea">UEA</span><div class="chips">{uea_html}</div></div>
      <div class="linha-u"><span class="lbl">Outras IES<br><small>{n_tot} PPGs · {n_ies} IES</small></span><div class="ies-l">{"".join(outras)}</div></div>
    </article>'''


def bloco_conf(r):
    itens = conferencia(r)
    ico = {'alerta': '⚠', 'nota': '≠', 'info': 'i'}
    lis = ''.join(f'<li class="cf-{k}"><span>{ico[k]}</span><p>{t}</p></li>' for k, t in itens)
    return f'<ul class="conf">{lis}</ul>', len([1 for k, _ in itens if k == 'alerta'])


def vista_rede(r):
    vr = por_rede[r]
    d_uea = distintos([v for v in vr if v['ies'] == 'UEA'])
    kpis = [(mil(val_rede[r]), 'previsto UEA · set/2027'), (str(len(IES_REDE[r])), 'universidades'),
            (str(len(AREAS[r])), 'áreas temáticas'), (str(len(distintos(vr))), f'PPGs distintos ({len(vr)} vinculações)'),
            (str(len(d_uea)), f'PPGs da UEA · {bol_rede[r]} bolsas')]
    kpi_html = ''.join(f'<div class="kpi"><b>{a}</b><span>{b}</span></div>' for a, b in kpis)
    conf, n_alerta = bloco_conf(r)
    modal = ''.join(f'<span class="md"><b>{mil(v)}</b>{e(m)}</span>' for m, v in sorted(modal_rede[r].items(), key=lambda x: -x[1]))
    return f'''
<section class="vista" id="v-{r}" data-vista="{r}" style="--c:{COR[r]};--cl:{CLARO[r]}">
  <header class="topo">
    <div class="logo-g"><img src="{logo(r)}" alt="{r}"></div>
    <div class="titulo">
      <span class="kick">CAPES Global · {e(NOME_REDE[r])} · coordenada pela {r}</span>
      <h1>{e(NOME_REDE[r])}: <em>quem está na rede</em> e onde a UEA entra</h1>
      <p>Anexo I da proposta (todas as áreas, IES e PPGs) cruzado com o comitê UEA e o balanço de setembro/2027.</p>
    </div>
  </header>
  <div class="kpis">{kpi_html}</div>

  <div class="grade g-mapa">
    <section class="card">
      <h2><span class="n">01</span>Universidades da rede</h2>
      {mapa(r)}
      <div class="leg-mapa"><span><i class="dot dot-uea"></i>UEA</span><span><i class="dot-lider" style="color:{COR[r]}">★</i>coordenadora</span><span><i class="dot" style="background:{COR[r]}"></i>associada</span><span>círculo ∝ nº de PPGs</span></div>
    </section>
    <section class="card">
      <h2><span class="n">02</span>Perfil das IES no Anexo I</h2>
      <p class="sub">PPGs distintos por código CAPES; nota média e perfil de notas dos PPGs de cada IES na rede.</p>
      {tabela_ies(r)}
      {LEG_NOTAS}
      <div class="mods"><h3>Previsto UEA por modalidade</h3><div class="mds">{modal}</div></div>
    </section>
  </div>

  <section class="card">
    <h2><span class="n">03</span>Matriz área temática × universidade</h2>
    <p class="sub">Nº de PPGs de cada IES em cada área. Passe o mouse para ver os programas; clique para abri-los no explorador. T# = temática do comitê UEA.</p>
    {matriz(r)}
  </section>

  <h2 class="secao"><span class="n">04</span>Áreas temáticas em detalhe</h2>
  <div class="areas">{"".join(card_area(r, a, i + 1) for i, a in enumerate(AREAS[r]))}</div>
  <div class="leg-chips">
    <span class="chip"><s>5</s><b>SIGLA</b><i>valor</i><u>n×</u></span> PPG da UEA · nota no Anexo I, valor previsto e nº de bolsas
    <span class="chip chip-0"><s>4</s><b>SIGLA</b></span> sem pedido em setembro
    <span class="chip chip-0"><s class="fora">!</s><b>SIGLA</b></span> pedido lançado em área onde o Anexo I não lista o PPG
  </div>

  <section class="card">
    <h2><span class="n">05</span>Conferência: comitê UEA × Anexo I {f'<span class="pill">{n_alerta} pontos a verificar</span>' if n_alerta else ''}</h2>
    {conf}
  </section>
</section>'''


def vista_geral():
    todos = len(vinc)
    ies_all = {u for r in REDES for u in IES_REDE[r]}
    uea_all = {}
    for v in vinc:
        if v['ies'] == 'UEA':
            uea_all.setdefault(v['cod'], v)
    kpis = [(mil(total_geral), 'previsto UEA · set/2027'), ('3', 'redes'), (str(len(ies_all)), 'universidades distintas'),
            (str(sum(len(AREAS[r]) for r in REDES)), 'áreas temáticas'), (str(todos), 'vinculações de PPGs'),
            (str(len(uea_all)), 'PPGs da UEA nos anexos')]
    kpi_html = ''.join(f'<div class="kpi"><b>{a}</b><span>{b}</span></div>' for a, b in kpis)
    cards = []
    nmax = max(STATS[r, u]['n'] for r in REDES for u in IES_REDE[r])
    for r in REDES:
        vr = por_rede[r]
        barras = ''.join(
            f'<div class="pb{" pb-uea" if u == "UEA" else ""}"><span>{e(u)}</span><div class="hbar"><span style="width:{STATS[r, u]["n"] / nmax * 100:.1f}%;'
            f'{"background:#0b6b3a" if u == "UEA" else ""}"></span></div><b>{STATS[r, u]["n"]}</b></div>' for u in IES_REDE[r])
        n_uea_areas = sum(1 for a in AREAS[r] if cel[r, a, 'UEA'])
        cards.append(f'''
      <button class="rede-card" data-ir="{r}" style="--c:{COR[r]};--cl:{CLARO[r]}">
        <div class="rc-top"><img src="{logo(r)}" alt="{r}"><div><b>{e(NOME_REDE[r])}</b><span>coordenada pela {r}</span></div></div>
        <div class="rc-val"><b>{mil(val_rede[r])}</b><span>previsto UEA · {bol_rede[r]} bolsas</span></div>
        <dl><div><dt>IES</dt><dd>{len(IES_REDE[r])}</dd></div><div><dt>Áreas</dt><dd>{len(AREAS[r])}</dd></div>
          <div><dt>PPGs</dt><dd>{len(distintos(vr))}</dd></div><div><dt>UEA em</dt><dd>{n_uea_areas}/{len(AREAS[r])} áreas</dd></div></dl>
        <h3>PPGs distintos por IES</h3>{barras}
        <span class="ir">Abrir rede →</span>
      </button>''')
    # UEA: PPG × rede
    linhas = []
    for cod, v0 in sorted(uea_all.items(), key=lambda kv: kv[1]['sigla']):
        cels = []
        for r in REDES:
            areas = sorted({v['area'] for v in vinc if v['rede'] == r and v['cod'] == cod and v['ies'] == 'UEA'})
            if not areas:
                cels.append('<td class="mudo c">·</td>')
                continue
            val = sum(x for (t, s), x in val_ppg.items() if s == v0['sigla'] and t.startswith(r + '-'))
            nb = sum(x for (t, s), x in bol_ppg.items() if s == v0['sigla'] and t.startswith(r + '-'))
            a_txt = '<br>'.join(e(a) for a in areas)
            cels.append(f'<td style="--c:{COR[r]}"><div class="ur">{f"<b>{mil(val)}</b> <u>{nb}×</u>" if val else "<b class=ad>sem pedido</b>"}'
                        f'<small>{a_txt}</small></div></td>')
        linhas.append(f'<tr><th><b>{e(v0["sigla"])}</b><span>{e(v0["ppg"])}</span></th><td class="c"><s class="nota-s">{v0["nota"]}</s></td>{"".join(cels)}</tr>')
    extra = sorted({p['sigla'] for p in estr['ppgs']} - {v['sigla'] for v in uea_all.values()})
    for sig in extra:  # cadastrados no comitê, ausentes dos anexos
        p = cad_sigla[sig]
        cels = []
        for r in REDES:
            if p['rede'] != r:
                cels.append('<td class="mudo c">·</td>')
                continue
            val = sum(x for (t, s), x in val_ppg.items() if s == sig and t.startswith(r + '-'))
            nb = sum(x for (t, s), x in bol_ppg.items() if s == sig and t.startswith(r + '-'))
            cels.append(f'<td style="--c:{COR[r]}"><div class="ur ur-fora">{f"<b>{mil(val)}</b> <u>{nb}×</u>" if val else "<b class=ad>sem pedido</b>"}'
                        f'<small>só no comitê T{p["temaId"][-1]} — fora do Anexo I</small></div></td>')
        linhas.append(f'<tr class="tr-fora"><th><b>{e(sig)}</b><span>{e(p["nome"])}</span></th><td class="c"><s class="nota-s">{p["notaCapes"]}</s></td>{"".join(cels)}</tr>')
    cab = ''.join(f'<th style="--c:{COR[r]}"><i class="dot" style="background:{COR[r]}"></i>{e(NOME_REDE[r])}</th>' for r in REDES)
    # notas por rede: IES coordenadora, associadas, UEA
    perf = []
    for r in REDES:
        d = distintos(por_rede[r])
        perf.append(f'<div class="pf"><span><i class="dot" style="background:{COR[r]}"></i>{e(NOME_REDE[r])} <small>({len(d)} PPGs)</small></span>{barra_notas(Counter(v["nota"] for v in d.values()))}</div>')
    perf.append(f'<div class="pf pf-uea"><span><i class="dot dot-uea"></i>UEA nos anexos <small>({len(uea_all)} PPGs)</small></span>{barra_notas(Counter(v["nota"] for v in uea_all.values()))}</div>')
    return f'''
<section class="vista on" id="v-geral" data-vista="geral" style="--c:#0b6b3a;--cl:#e3f1e9">
  <header class="topo">
    <div class="logo-g"><img src="{logo('UEA')}" alt="UEA"></div>
    <div class="titulo">
      <span class="kick">CAPES Global.EDU · PROPESP / UEA · mapeamento dos Anexos I</span>
      <h1>A UEA nas <em>3 redes CAPES Global</em></h1>
      <p>Universidades, áreas temáticas e PPGs de cada rede, com os valores previstos para a UEA em setembro/2027.</p>
    </div>
  </header>
  <div class="kpis">{kpi_html}</div>
  <div class="redes-cards">{"".join(cards)}</div>
  <section class="card">
    <h2><span class="n">01</span>PPGs da UEA em cada rede</h2>
    <p class="sub">Área(s) do Anexo I em que cada PPG aparece e o valor previsto em setembro/2027. Linhas tracejadas: PPGs cadastrados no comitê UEA que não constam nos anexos.</p>
    <div class="rolar"><table class="t-uea"><thead><tr><th>PPG</th><th class="c">Nota</th>{cab}</tr></thead><tbody>{"".join(linhas)}</tbody></table></div>
  </section>
  <section class="card">
    <h2><span class="n">02</span>Perfil de notas CAPES</h2>
    <p class="sub">Distribuição das notas dos PPGs distintos listados em cada rede, comparada aos PPGs da UEA.</p>
    <div class="pfs">{"".join(perf)}</div>
    {LEG_NOTAS}
  </section>
</section>'''


def vista_explorar():
    return '''
<section class="vista" id="v-explorar" data-vista="explorar" style="--c:#14213d;--cl:#e8ebf2">
  <header class="topo topo-s">
    <div class="titulo">
      <span class="kick">Explorador · todas as vinculações dos Anexos I</span>
      <h1>Explorar <em>PPGs</em></h1>
    </div>
  </header>
  <section class="card">
    <div class="filtros">
      <label>Rede<select id="f-rede"><option value="">Todas</option></select></label>
      <label>Área temática<select id="f-area"><option value="">Todas</option></select></label>
      <label>Universidade<select id="f-ies"><option value="">Todas</option></select></label>
      <label>Nota mínima<select id="f-nota"><option value="0">Qualquer</option><option>3</option><option>4</option><option>5</option><option>6</option><option>7</option></select></label>
      <label class="busca">Buscar PPG ou código<input id="f-q" type="search" placeholder="ex.: biotecnologia, 12008010…"></label>
      <button id="f-limpar" class="bt">Limpar</button>
      <button id="f-csv" class="bt bt-pri">Baixar CSV</button>
    </div>
    <p class="sub" id="f-resumo"></p>
    <div class="rolar t-wrap"><table class="t-exp"><thead><tr>
      <th data-k="rede">Rede</th><th data-k="area">Área temática</th><th data-k="ies">IES</th><th data-k="ppg">Programa</th>
      <th data-k="cod">Código CAPES</th><th data-k="nota" class="c">Nota</th><th data-k="pag" class="c">Pág.</th></tr></thead><tbody id="f-corpo"></tbody></table></div>
  </section>
</section>'''


# ---------------------------------------------------------------- página
abas = ('<button class="aba on" data-vista="geral"><img src="' + logo('UEA') + '" alt="">Visão geral</button>'
        + ''.join(f'<button class="aba" data-vista="{r}" style="--c:{COR[r]}"><img src="{logo(r)}" alt="">{e(NOME_REDE[r])}</button>' for r in REDES)
        + '<button class="aba" data-vista="explorar">🔍 Explorar PPGs</button>')
DADOS = {'redes': {r: NOME_REDE[r] for r in REDES}, 'cor': COR, 'areas': AREAS,
         'ies': IES_REDE, 'v': [[v['rede'], v['area'], v['ies'], v['ppg'], v['cod'], v['nota'], v['pag']] for v in vinc]}
CSS = (AQUI / 'dashboard.css').read_text(encoding='utf-8')
JS = (AQUI / 'dashboard.js').read_text(encoding='utf-8')
LIB = (VENDOR / 'html-to-image.js').read_text(encoding='utf-8')

pagina = f'''<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>UEA nas Redes CAPES Global</title>
<style>{CSS}</style>
</head>
<body>
<nav class="barra">
  <span class="barra-t">UEA nas Redes CAPES Global<small>Anexos I · balanço setembro/2027</small></span>
  <div class="abas" role="tablist">{abas}</div>
  <div class="acoes"><button data-acao="png" class="pri">PNG desta vista</button></div>
</nav>
<main id="palco">{vista_geral()}{"".join(vista_rede(r) for r in REDES)}{vista_explorar()}
<footer class="fonte">Fontes: CAPES_Global_Mapeamento_3_Redes_PPGs_IES.xlsm (Anexos I das propostas) · Balanço CAPES Global 2027 — Setembro | Bolsas a atender · cadastro do comitê CAPES Global (Gerenciador PROPESP).</footer>
</main>
<div id="tip" role="tooltip"></div>
<script>{LIB}</script>
<script>const DADOS = {json.dumps(DADOS, ensure_ascii=False)};
{JS}</script>
</body>
</html>'''

for saida in ('dashboard_redes.html', 'index.html'):  # index.html: GitHub Pages
    (AQUI / saida).write_text(pagina, encoding='utf-8')
print('ok', len(pagina) // 1024, 'KB')
for r in REDES:
    print(r, mil(val_rede[r]), len(por_rede[r]), 'vinculações', len(distintos(por_rede[r])), 'PPGs')
    for k, t in conferencia(r):
        print('  ', k, re.sub('<[^>]+>', '', t))
