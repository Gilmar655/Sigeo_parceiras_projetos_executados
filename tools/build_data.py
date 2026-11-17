"""Gera site/data/base.js a partir das planilhas (Start, Conecta, Engelmig, Projetos, Equipamentos)."""
import sys, json, re, datetime as dt
from collections import Counter, defaultdict
import pandas as pd

D = sys.argv[1]; OUT = sys.argv[2]
files = {
    'Start': f'{D}/985afa8a-Start_16_11_2026.xlsx',
    'Conecta': f'{D}/c6bfa2a6-Conecta_16_11_2026.xlsx',
    'Engelmig': f'{D}/adac6799-Engelmig_16_11.xlsx',
}
proj_file = f'{D}/f1b61bd5-Projetos_01_01_31_maio.xlsx'
eq_file = f'{D}/b029e26b-Equipamentos_Buscape_Links.xlsx'

def s(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NaT: return ''
    return re.sub(r'\s+', ' ', str(v)).strip()

def d(v):
    if v is None or v is pd.NaT or (isinstance(v, float) and pd.isna(v)): return ''
    if isinstance(v, (dt.datetime, pd.Timestamp)): return v.strftime('%Y-%m-%d')
    return s(v)[:10]

def hm(v):
    if v is None or v is pd.NaT or (isinstance(v, float) and pd.isna(v)): return ''
    if isinstance(v, (dt.datetime, pd.Timestamp)): return v.strftime('%H:%M')
    if isinstance(v, dt.time): return v.strftime('%H:%M')
    m = re.search(r'(\d{1,2}):(\d{2})', str(v)); return f'{int(m[1]):02d}:{m[2]}' if m else ''

def num(v):
    try:
        f = float(v); return 0 if pd.isna(f) else (int(f) if f.is_integer() else round(f, 2))
    except Exception: return 0

rows = []
for origem, f in files.items():
    df = pd.read_excel(f, dtype={'Numero PowerON': str, 'Equipamentos': str, 'Projeto': str})
    for line, r in enumerate(df.to_dict('records'), start=2):
        rows.append([
            s(r['Projeto']).upper(), d(r['Data Programacao']),
            s(r['Status Programacao']), s(r['Tipo Intervencao']), s(r['Numero PowerON']),
            s(r['Equipamentos']).upper(), num(r['CHI']), hm(r['Horario Inicio']), hm(r['Horario Fim']),
            s(r['Contratada']), origem, line,
        ])
print('programacoes lidas', len(rows))

# Regras de exclusão (escopo do painel)
EXCLUIR = {('Start', 'Manutenção de Redes e Linhas Aéreas'), ('Conecta', 'Poda de Árvores')}
_pr = pd.read_excel(proj_file, dtype=str, usecols=['Projeto', 'Descrição do Projeto'])
_desc = {s(a).upper(): s(b) for a, b in zip(_pr['Projeto'], _pr['Descrição do Projeto'])}
rows = [r for r in rows if (r[10], _desc.get(r[0], '')) not in EXCLUIR]
print('programacoes apos exclusoes', len(rows))

# Base de projetos (cadastro)
needed = {r[0] for r in rows}
pr = pd.read_excel(proj_file, dtype=str)
projects = {}
for r in pr.itertuples(index=False):
    code = s(r[0]).upper()
    if code not in needed: continue
    projects[code] = [s(r[1]), s(r[2]), s(r[3]), s(r[4]), s(r[5]).upper(), s(r[6]), s(r[7]), s(r[8]),
                      s(r[9]), s(r[10]), d(r[11]), s(r[12])]
print('projetos cadastro', len(projects), 'de', len(needed))

# Coordenadas por equipamento
eq = pd.read_excel(eq_file, sheet_name='Equipamentos', dtype={'Número do Equipamento': str, 'Circuito': str})
eqmap = {}
for code, circ, lat, lng in zip(eq['Número do Equipamento'], eq['Circuito'], eq['Latitude'], eq['Longitude']):
    if pd.isna(lat) or pd.isna(lng): continue
    eqmap[s(code).upper()] = [round(float(lat), 6), round(float(lng), 6), s(circ)]
tok = lambda e: [t for t in re.split(r'[,;|\s]+', e) if t]
coords = {}
for r in rows:
    for t in tok(r[5]):
        if t in eqmap: coords[t] = eqmap[t]
print('equipamentos com coordenada usados', len(coords))

dates = sorted(r[1] for r in rows if r[1])
meta = {
    'generated': dt.datetime.now().strftime('%Y-%m-%dT%H:%M'),
    'periodo': [dates[0], dates[-1]],
    'fontes': {k: sum(1 for r in rows if r[10] == k) for k in files},
    'arquivos': ['Start_16_11_2026.xlsx', 'Conecta_16_11_2026.xlsx', 'Engelmig_16_11.xlsx',
                 'Projetos_01_01_31_maio.xlsx', 'Equipamentos_Buscape_Links.xlsx (base 12/06/2026)'],
}
data = {
    'meta': meta,
    'progCols': ['Projeto', 'Data Programacao', 'Status Programacao', 'Tipo Intervencao', 'Numero PowerON',
                 'Equipamentos', 'CHI', 'Horario Inicio', 'Horario Fim', 'Contratada', 'Parceira', 'Linha'],
    'prog': rows,
    'projCols': ['Nota', 'Empresa Executante', 'Descrição do Projeto', 'Rua', 'Regiao', 'Pendente', 'UPS Orçada',
                 'Status do Projeto', 'Priorizado', 'Prioritario', 'Data da Priorizacao', 'Motivo da Priorizacao'],
    'projects': projects,
    'coords': coords,
}
with open(OUT, 'w', encoding='utf-8') as fh:
    fh.write('window.SIGEO_DATA=')
    json.dump(data, fh, ensure_ascii=False, separators=(',', ':'))
    fh.write(';\n')
print('ok')
