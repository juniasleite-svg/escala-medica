"""
Gerador de Excel — formato CM5 exato
Reproduz as 8 abas do arquivo Escala_ClinicaMedica_5ANO_Grupo_A_T6_vJUNIA
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date, timedelta
import io
import json

# ── Paleta de cores exata do CM5 ─────────────────────────────────────────────
C = {
    # Headers
    "H1":  "1F4E79",  # azul escuro (título)
    "H2":  "2E75B6",  # azul médio (sub-header)
    "H3":  "D6E4F0",  # azul claro (header datas)
    "FDS": "C5CAD6",  # cinza (fim de semana)
    # SGs - cor da linha de nome
    "SG1": "D6E4F7", "SG2": "D4EED4", "SG3": "FFF3CD",
    "SG4": "FFD6D6", "SG5": "EDD9F5", "SG6": "D9F5E8",
    "SG7": "FDE8D8", "SG8": "E8F8E8",
    # Locais - cor das células de dados
    "ENF": "AED6F1",  # azul claro (enfermaria)
    "PS":  "A9DFBF",  # verde claro (PS)
    "AMB": "FAD7A0",  # laranja claro (ambulatório)
    "LOC4":"D7BDE2",  # lilás
    "LOC5":"F9E79F",  # amarelo
    "LOC6":"FADBD8",  # rosa
    # Especiais
    "FDS_CELL": "F5F5F5",  # fim de semana sem plantão
    "FDS_PLT":  "D7BDE2",  # plantão FDS / PA★
    "TARDE_R":  "FFF2CC",  # tarde reduzida
    "VERDE":    "E2EFDA",  # área verde
    "TOTAL":    "FFF2CC",  # total horas
    "CIND":     "FFFFFF",
    "ENF_FDS_H":"EBF5FB",  # header plantões enf fds
    "PS_FDS_H": "EAFAF1",  # header plantões ps fds
    # Alternância linhas
    "ALT1": "EEF3FB",
    "ALT2": "FFFFFF",
    "SEP":  "2E75B6",  # separador de SG
}

DIAS_PT = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]

def _normalizar_data(data_str, ano="2026"):
    """Normaliza data para formato DD/MM/YYYY."""
    if not data_str: return ""
    s = str(data_str).strip()
    # Já está completo
    if len(s) == 10 and "/" in s: return s  # DD/MM/YYYY
    if len(s) == 10 and "-" in s:  # YYYY-MM-DD
        parts = s.split("-")
        return f"{parts[2]}/{parts[1]}/{parts[0]}"
    # DD/MM — adicionar ano
    if len(s) == 5 and "/" in s:
        return f"{s}/{ano}"
    if len(s) == 5 and "-" in s:
        parts = s.split("-")
        return f"{parts[1]}/{parts[0]}/{ano}"
    return s



def _cor(rgb):
    return PatternFill("solid", fgColor=rgb)

def _borda(color="BFBFBF", style="thin"):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _cel(ws, row, col, val="", bold=False, bg=None, fc="000000",
         halign="center", valign="center", wrap=False, sz=9, italic=False,
         border=True, border_color="BFBFBF"):
    c = ws.cell(row=row, column=col, value=val)
    c.font = Font(bold=bold, color=fc, size=sz, italic=italic)
    c.alignment = Alignment(horizontal=halign, vertical=valign, wrap_text=wrap)
    if border:
        c.border = _borda(border_color)
    if bg:
        c.fill = _cor(bg)
    return c

def _header(ws, row, col, val, bg=C["H1"], fc="FFFFFF", sz=11, bold=True, span=None):
    c = _cel(ws, row, col, val, bold=bold, bg=bg, fc=fc, sz=sz)
    if span and span > 1:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+span-1)
    ws.row_dimensions[row].height = 22
    return c

def _cor_local(local_nome, locais_cfg):
    """Retorna cor da célula baseada no local."""
    cores = [C["ENF"], C["PS"], C["AMB"], C["LOC4"], C["LOC5"], C["LOC6"]]
    nomes = [l.get("nome","").lower() for l in locais_cfg]
    nome_lower = local_nome.lower()
    for i, n in enumerate(nomes):
        if n and (n in nome_lower or nome_lower in n):
            return cores[i % len(cores)]
    return C["ENF"]

def _cor_sg(sg_num):
    key = f"SG{sg_num}"
    return C.get(key, C["SG1"])

def _gerar_datas(data_inicio_str, num_semanas):
    """Retorna lista de semanas, cada uma com 7 datas."""
    try:
        d0 = date.fromisoformat(data_inicio_str)
    except:
        d0 = date.today()
    return [[d0 + timedelta(weeks=s, days=d) for d in range(7)] for s in range(num_semanas)]

def _turno_compacto(entrada):
    """Converte entrada da IA para formato compacto: Enf(M+T)"""
    if not entrada or entrada in ("—", "-", "", "None"):
        return "—"
    return str(entrada)

def gerar_excel_completo(dados, config):
    """
    dados: dict com calendario_rodizio, escala_detalhada, resumo_horas, auditoria
    config: dict com especialidade, grupo, turma, data_inicio, num_semanas, locais, alunos_por_sg
    """
    esp = config.get("especialidade","")
    grupo = config.get("grupo","")
    turma = config.get("turma","")
    ano = config.get("ano_curso","")
    titulo = f"{esp} — {ano} — {grupo} / {turma}"
    
    locais_cfg = config.get("locais", [])
    alunos_por_sg = config.get("alunos_por_sg", {})
    num_semanas = int(config.get("num_semanas", 8))
    data_inicio = config.get("data_inicio", str(date.today()))
    rodizio_desc = config.get("rodizio_desc", "")
    regras = config.get("regras_especiais", {})
    
    semanas = _gerar_datas(data_inicio, num_semanas)
    
    # Detectar ano da escala
    ano_escala = "2026"
    try:
        ano_escala = date.fromisoformat(data_inicio).strftime("%Y")
    except: pass

    # Montar índice de escala detalhada por aluno/data
    escala_det = dados.get("escala_detalhada", [])
    # {nome: {data_normalizada: "turno_compacto"}}
    escala_por_aluno = {}
    for entry in escala_det:
        alunos = entry.get("alunos", [])
        if isinstance(entry.get("nome"), str) and entry.get("nome"):
            alunos = [entry["nome"]]
        data_raw = str(entry.get("data",""))
        data_str = _normalizar_data(data_raw, ano_escala)
        turno = entry.get("turno","")
        local = entry.get("local","")
        abrev = _abrev_local(local, locais_cfg)
        codigo = _codigo_turno(turno, local, locais_cfg)
        valor = f"{abrev}({codigo})" if codigo else f"{abrev}"
        for nome in alunos:
            if nome not in escala_por_aluno:
                escala_por_aluno[nome] = {}
            if data_str in escala_por_aluno[nome]:
                existing = escala_por_aluno[nome][data_str]
                if valor not in existing:
                    escala_por_aluno[nome][data_str] = _combinar(existing, valor)
            else:
                escala_por_aluno[nome][data_str] = valor

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _aba_resumo_geral(wb, titulo, config, dados, semanas)
    _aba_alunos(wb, titulo, alunos_por_sg, config)
    _aba_calendario_rodizio(wb, titulo, dados.get("calendario_rodizio",[]), config, semanas)
    _aba_escala_nominal(wb, titulo, escala_det, config, semanas, locais_cfg, ano_escala)
    _aba_resumo_horas(wb, titulo, dados.get("resumo_horas",[]), config, semanas, locais_cfg)
    _aba_regras(wb, titulo, config, dados)
    _aba_escala_subgrupo(wb, titulo, alunos_por_sg, escala_por_aluno, config, semanas, locais_cfg)
    _aba_escala_individual(wb, titulo, alunos_por_sg, escala_por_aluno, config, semanas, locais_cfg)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()


def _abrev_local(local_nome, locais_cfg):
    """Retorna abreviação do local."""
    for l in locais_cfg:
        n = l.get("nome","")
        if n and (n.lower() in local_nome.lower() or local_nome.lower() in n.lower()):
            ab = l.get("abrev","")
            return ab if ab else n[:3]
    # Fallback: primeiras 3 letras
    return local_nome[:3] if local_nome else "?"

def _codigo_turno(turno, local, locais_cfg):
    """Converte turno para código compacto."""
    t = turno.lower() if turno else ""
    if "fds" in t or "★" in turno or "pa" in t.lower():
        return "FDS"
    if "cinderela" in t or t == "c": return "C"
    if "reduz" in t or t == "r": return "R"
    if "manhã" in t or t == "m" or "manha" in t: return "M"
    if "tarde" in t or t == "t": return "T"
    if "noite" in t or t == "n": return "N"
    if "+" in t: return t.upper()
    return t[:1].upper() if t else "M"

def _combinar(v1, v2):
    """Combina dois valores de turno."""
    if v1 == v2: return v1
    # Extrair local e turno
    def parse(v):
        if "(" in v and ")" in v:
            loc = v[:v.index("(")]
            t = v[v.index("(")+1:v.index(")")]
            return loc, t
        return v, ""
    loc1, t1 = parse(v1)
    loc2, t2 = parse(v2)
    if loc1 == loc2:
        # Mesmo local, combinar turnos
        turnos = set()
        for t in [t1, t2]:
            for part in t.split("+"):
                turnos.add(part)
        ordem = ["M","T","R","C","F","N","FDS"]
        combined = "+".join([x for x in ordem if x in turnos])
        return f"{loc1}({combined})" if combined else loc1
    return f"{v1}/{v2}"


# ── ABA 1: RESUMO GERAL ───────────────────────────────────────────────────────
def _aba_resumo_geral(wb, titulo, config, dados, semanas):
    ws = wb.create_sheet("Resumo Geral")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 55

    esp = config.get("especialidade","")
    grupo = config.get("grupo","")
    turma = config.get("turma","")
    ano = config.get("ano_curso","")
    locais_cfg = config.get("locais",[])
    alunos_por_sg = config.get("alunos_por_sg",{})
    regras = config.get("regras_especiais",{})
    rodizio = config.get("rodizio_desc","")

    d_ini = semanas[0][0].strftime("%d/%m/%Y") if semanas else ""
    d_fim = semanas[-1][-1].strftime("%d/%m/%Y") if semanas else ""
    n_sem = len(semanas)
    n_alunos = sum(len(v) for v in alunos_por_sg.values())

    _header(ws, 1, 1, f"{esp} — {ano} — {grupo} / {turma}", span=2, sz=13)
    _header(ws, 2, 1, f"Escala de Estágio  |  {d_ini} – {d_fim}  |  {n_sem} semanas  |  Rodízio {rodizio[:50]}", span=2, bg=C["H2"], sz=10)
    ws.row_dimensions[2].height = 16

    dados_resumo = [
        ("Período", f"{d_ini} a {d_fim} ({n_sem} semanas)"),
        ("Grupo / Turma", f"{grupo} — {turma} — {ano}"),
        ("Total de alunos", str(n_alunos)),
        ("Subgrupos", " | ".join([f"SG{k}({len(v)}al)" for k,v in sorted(alunos_por_sg.items(), key=lambda x:int(x[0]))])),
        ("Locais", " · ".join([l.get("nome","") for l in locais_cfg])),
        ("Rodízio", rodizio),
    ]
    # Regras de cada local
    for l in locais_cfg:
        linha_m = l.get("manha","")
        linha_t = l.get("tarde","")
        dados_resumo.append((l.get("nome",""), f"Manhã: {linha_m} | Tarde: {linha_t}"))
    dados_resumo += [
        ("Quinta-feira", regras.get("quinta","ENAMED — sem campo tarde")),
        ("Terça-feira", regras.get("terca","Aula — campo tarde encurtado")),
        ("Limite CH", f"{regras.get('limite_ch',40)}h/sem | Absoluto: {regras.get('limite_abs',43)}h"),
        ("FDS", regras.get("fds","")),
    ]

    row = 4
    for label, val in dados_resumo:
        if not val: continue
        _cel(ws, row, 1, label, bold=True, bg=C["H3"], halign="left", sz=9)
        _cel(ws, row, 2, val, halign="left", sz=9, wrap=True)
        ws.row_dimensions[row].height = 15
        row += 1

    # Legenda
    row += 1
    _header(ws, row, 1, "LEGENDA DE LOCAIS", span=2, bg=C["H2"])
    row += 1
    cores_loc = [C["ENF"], C["PS"], C["AMB"], C["LOC4"], C["LOC5"], C["LOC6"]]
    for i, l in enumerate(locais_cfg):
        cor = cores_loc[i % len(cores_loc)]
        ab = l.get("abrev", l.get("nome","")[:3])
        _cel(ws, row, 1, ab, bold=True, bg=cor, sz=9)
        _cel(ws, row, 2, l.get("nome",""), halign="left", bg=cor, sz=9)
        row += 1

    # Legenda de turnos
    row += 1
    _header(ws, row, 1, "LEGENDA DE TURNOS", span=2, bg=C["H2"])
    row += 1
    turnos_leg = [
        ("M", "Manhã (horário conforme local)", None),
        ("T", "Tarde completa", None),
        ("R", "Tarde reduzida (fundo amarelo)", C["TARDE_R"]),
        ("F", "Plantão FDS manhã Enf/PS", C["FDS"]),
        ("★ / FDS", "Plantão PA/FDS — fundo lilás", C["FDS_PLT"]),
        ("AV", "Área Verde — tarde livre", C["VERDE"]),
        ("C", "Cinderela", None),
        ("—", "Sem atividade (FDS sem plantão)", C["FDS_CELL"]),
    ]
    for ab, desc, cor in turnos_leg:
        _cel(ws, row, 1, ab, bold=True, bg=cor or "FFFFFF", sz=9)
        _cel(ws, row, 2, desc, halign="left", bg=cor or "FFFFFF", sz=9)
        row += 1

    # Auditoria
    audit = dados.get("auditoria",{})
    if audit:
        row += 1
        _header(ws, row, 1, "AUDITORIA", span=2, bg=C["H1"])
        row += 1
        status = "✅ APROVADA" if audit.get("aprovado") else "❌ COM ERROS"
        cor_st = "C6EFCE" if audit.get("aprovado") else "FFC7CE"
        _cel(ws, row, 1, "Status", bold=True, bg="F2F2F2", sz=9)
        _cel(ws, row, 2, status, bold=True, bg=cor_st, sz=9, halign="left")
        for err in audit.get("erros",[]):
            row += 1
            _cel(ws, row, 1, "ERRO", bold=True, bg="FFC7CE", sz=9)
            _cel(ws, row, 2, err, bg="FFC7CE", sz=9, halign="left", wrap=True)
        for av in audit.get("avisos",[]):
            row += 1
            _cel(ws, row, 1, "AVISO", bold=True, bg="FFEB9C", sz=9)
            _cel(ws, row, 2, av, bg="FFEB9C", sz=9, halign="left", wrap=True)


# ── ABA 2: ALUNOS ────────────────────────────────────────────────────────────
def _aba_alunos(wb, titulo, alunos_por_sg, config):
    ws = wb.create_sheet("Alunos")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 5
    ws.column_dimensions["C"].width = 38
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 8
    ws.column_dimensions["F"].width = 10

    turma = config.get("turma","")
    grupo = config.get("grupo","")
    ano = config.get("ano_curso","")

    _header(ws, 1, 1, f"LISTA DE ALUNOS — {grupo} / {turma} — {ano}", span=6, sz=11)
    for i, h in enumerate(["Nº","SG","Nome Completo","RA","Turma","Par"],1):
        _cel(ws, 2, i, h, bold=True, bg=C["H2"], fc="FFFFFF", sz=9)

    row = 3; num = 1
    for sg_key in sorted(alunos_por_sg.keys(), key=lambda x: int(x[0]) if x[0].isdigit() else 0):
        sg_num = int(sg_key) if sg_key.isdigit() else 1
        cor = _cor_sg(sg_num)
        for i_al, nome in enumerate(alunos_por_sg[sg_key]):
            cor_row = C["ALT1"] if i_al % 2 == 0 else C["ALT2"]
            _cel(ws, row, 1, num, bg=cor_row, sz=9)
            _cel(ws, row, 2, f"SG{sg_num}", bold=True, bg=cor, sz=9)
            _cel(ws, row, 3, nome, halign="left", bg=cor_row, sz=9)
            _cel(ws, row, 4, "", bg=cor_row, sz=9)  # RA
            _cel(ws, row, 5, turma, bg=cor_row, sz=9)
            _cel(ws, row, 6, "", bg=cor_row, sz=9)  # Par
            num += 1; row += 1


# ── ABA 3: CALENDÁRIO DE RODÍZIO ─────────────────────────────────────────────
def _aba_calendario_rodizio(wb, titulo, calendario, config, semanas):
    ws = wb.create_sheet("Calendário de Rodízio")
    ws.sheet_view.showGridLines = False

    grupo = config.get("grupo","")
    turma = config.get("turma","")
    rodizio_desc = config.get("rodizio_desc","")
    locais_cfg = config.get("locais",[])
    alunos_por_sg = config.get("alunos_por_sg",{})

    sgs = sorted(alunos_por_sg.keys(), key=lambda x: int(x[0]) if x[0].isdigit() else 0)
    n_sgs = len(sgs)

    _header(ws, 1, 1, f"CALENDÁRIO DE RODÍZIO — {grupo} / {turma}", span=2+n_sgs, sz=11)
    _header(ws, 2, 1, rodizio_desc[:120], span=2+n_sgs, bg=C["H2"], sz=9)

    # Headers SGs
    _cel(ws, 3, 1, "Par / SG", bold=True, bg=C["H2"], fc="FFFFFF", sz=9)
    _cel(ws, 3, 2, "Período", bold=True, bg=C["H2"], fc="FFFFFF", sz=9)
    for i, sg in enumerate(sgs):
        n_al = len(alunos_por_sg.get(sg,[]))
        _cel(ws, 3, 3+i, f"SG{sg}({n_al})", bold=True, bg=C["H2"], fc="FFFFFF", sz=9)

    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 22
    for i in range(n_sgs):
        ws.column_dimensions[get_column_letter(3+i)].width = 22

    cores_loc = [C["ENF"], C["PS"], C["AMB"], C["LOC4"], C["LOC5"], C["LOC6"]]
    loc_nomes = [l.get("nome","") for l in locais_cfg]
    loc_cor_map = {n: cores_loc[i%len(cores_loc)] for i,n in enumerate(loc_nomes)}

    for ri, sem in enumerate(calendario):
        row = ri + 4
        sem_num = sem.get("semana", ri+1)
        periodo = sem.get("periodo","")
        if not periodo and ri < len(semanas):
            d0 = semanas[ri][0]; d1 = semanas[ri][-1]
            periodo = f"{d0.strftime('%d/%m')}–{d1.strftime('%d/%m/%Y')}"
        _cel(ws, row, 1, f"Sem {sem_num}", bold=True, bg="F2F2F2", sz=9)
        _cel(ws, row, 2, periodo, bg="F2F2F2", sz=9)
        aloc = sem.get("alocacao",{})
        for i, sg in enumerate(sgs):
            local = aloc.get(f"SG{sg}", aloc.get(sg, ""))
            cor = next((c for n,c in loc_cor_map.items() if n.lower() in local.lower()), "F2F2F2") if local else "F2F2F2"
            _cel(ws, row, 3+i, local, bg=cor, sz=9)

    # Blocos
    row = len(calendario) + 5
    _header(ws, row, 1, "BLOCOS DE RODÍZIO", span=2+n_sgs, bg=C["H2"])
    for bl in config.get("blocos",[]):
        row += 1
        _cel(ws, row, 1, bl.get("nome",""), bold=True, bg="F2F2F2", sz=9)
        _cel(ws, row, 2, bl.get("periodo",""), bg="F2F2F2", sz=9)
        _cel(ws, row, 3, bl.get("descricao",""), halign="left", bg="F2F2F2", sz=9, wrap=True)
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=2+n_sgs)


# ── ABA 4: ESCALA NOMINAL DETALHADA ──────────────────────────────────────────
def _aba_escala_nominal(wb, titulo, escala_det, config, semanas, locais_cfg, ano_escala="2026"):
    ws = wb.create_sheet("Escala Nominal Detalhada")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "E4"

    grupo = config.get("grupo","")
    turma = config.get("turma","")
    d_ini = semanas[0][0].strftime("%d/%m/%Y") if semanas else ""
    d_fim = semanas[-1][-1].strftime("%d/%m/%Y") if semanas else ""
    regras = config.get("regras_especiais",{})

    subtitulo = f"Terça = {regras.get('terca','12-16h')}  ·  Quinta = {regras.get('quinta','ENAMED')}"

    _header(ws, 1, 1, f"ESCALA NOMINAL DETALHADA — {grupo} / {turma}  |  {d_ini}–{d_fim}  |  rodízio equilibrado ✓",
            span=10, sz=11)
    _header(ws, 2, 1, subtitulo, span=10, bg=C["H2"], sz=9)

    for i, h in enumerate(["Sem","Data","Dia","Local","Turno","Horário","h","SG","Nome","RA"],1):
        _cel(ws, 3, i, h, bold=True, bg=C["H2"], fc="FFFFFF", sz=9)

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 5
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 5
    ws.column_dimensions["I"].width = 36
    ws.column_dimensions["J"].width = 14

    sem_atual = None
    for entry in escala_det:
        sem = entry.get("semana","")
        if sem != sem_atual:
            sem_atual = sem
            # Linha separadora de semana
            periodo = entry.get("data","")
            label = f"SEMANA {sem}"
            if semanas and isinstance(sem, int) and 1 <= sem <= len(semanas):
                d0 = semanas[sem-1][0]; d1 = semanas[sem-1][-1]
                label = f"SEMANA {sem}  —  {d0.strftime('%d/%m')} a {d1.strftime('%d/%m/%Y')}"
            row = ws.max_row + 1
            _header(ws, row, 1, label, span=10, bg=C["H2"], sz=9)

        row = ws.max_row + 1
        local = entry.get("local","")
        turno = entry.get("turno","")
        data_raw = str(entry.get("data",""))
        data_norm = _normalizar_data(data_raw, ano_escala if 'ano_escala' in dir() else "2026")
        cor = _cor_local(local, locais_cfg)

        # Cor especial para FDS e cinderela
        turno_lower = turno.lower() if turno else ""
        if "fds" in turno_lower or "★" in turno or "plantão" in turno_lower:
            cor = C["FDS_PLT"]
        elif "reduz" in turno_lower or turno.upper() == "R":
            cor = C["TARDE_R"]
        elif "verde" in turno_lower or turno.upper() == "AV":
            cor = C["VERDE"]

        alunos = entry.get("alunos",[])
        nome_str = entry.get("nome","") if isinstance(entry.get("nome"), str) else ""
        if alunos and not nome_str:
            nome_str = " | ".join(alunos)

        vals = [sem, data_norm, entry.get("dia",""), local,
                turno, entry.get("horario",""), entry.get("horas",""),
                entry.get("sg",""), nome_str, entry.get("ra","")]
        for ci, v in enumerate(vals, 1):
            _cel(ws, row, ci, v, bg=cor, sz=8,
                 halign="left" if ci in [4,9] else "center", wrap=(ci==9))
        ws.row_dimensions[row].height = 13


# ── ABA 5: RESUMO DE HORAS ───────────────────────────────────────────────────
def _aba_resumo_horas(wb, titulo, resumo_horas, config, semanas, locais_cfg):
    ws = wb.create_sheet("Resumo de Horas")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "D3"

    grupo = config.get("grupo","")
    turma = config.get("turma","")
    n_sem = len(semanas)

    _header(ws, 1, 1, f"RESUMO DE HORAS — {grupo} / {turma}", span=5+n_sem, sz=11)

    # Headers semanas com data
    sem_labels = []
    for i, sem in enumerate(semanas):
        d = sem[0].strftime("%d/%m")
        sem_labels.append(f"S{i+1}\n{d}")

    headers = ["SG","Nome","RA"] + sem_labels + ["TOTAL","Cind","Plantões\nEnf FDS","Plantões\nPS FDS"]
    for i, h in enumerate(headers, 1):
        bg = C["H1"] if h in ["SG","TOTAL"] else (C["ENF_FDS_H"] if "Enf" in h else (C["PS_FDS_H"] if "PS" in h else C["H2"]))
        fc = "FFFFFF" if bg in [C["H1"],C["H2"]] else "000000"
        _cel(ws, 2, i, h, bold=True, bg=bg, fc=fc, sz=9, wrap=True)
        ws.row_dimensions[2].height = 28

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 14
    for i in range(n_sem):
        ws.column_dimensions[get_column_letter(4+i)].width = 8
    ws.column_dimensions[get_column_letter(4+n_sem)].width = 10
    for j in range(3):
        ws.column_dimensions[get_column_letter(5+n_sem+j)].width = 12

    alunos_por_sg = config.get("alunos_por_sg",{})
    # Montar mapa de resumo_horas por nome
    rh_map = {r.get("nome",""): r for r in resumo_horas} if resumo_horas else {}

    row = 3
    for sg_key in sorted(alunos_por_sg.keys(), key=lambda x: int(x[0]) if x[0].isdigit() else 0):
        sg_num = int(sg_key) if sg_key.isdigit() else 1
        cor_sg = _cor_sg(sg_num)
        # Cores dos blocos (local) para as semanas
        for i_al, nome in enumerate(alunos_por_sg[sg_key]):
            cor_row = C["ALT1"] if i_al % 2 == 0 else C["ALT2"]
            rh = rh_map.get(nome, {})
            semanas_h = rh.get("semanas",[0]*n_sem)
            total = rh.get("total_horas", sum(semanas_h) if semanas_h else 0)

            _cel(ws, row, 1, f"SG{sg_num}", bold=True, bg=cor_sg, sz=9)
            _cel(ws, row, 2, nome, halign="left", bg=cor_row, sz=9)
            _cel(ws, row, 3, str(rh.get("ra","")), bg=cor_row, sz=9)

            # Horas por semana com cor do local
            for i_sem in range(n_sem):
                h = semanas_h[i_sem] if i_sem < len(semanas_h) else 0
                # Cor da célula baseada no local daquela semana
                cor_sem = _cor_semana_aluno(nome, i_sem+1, resumo_horas, locais_cfg)
                cor_h = "FFC7CE" if h > 43 else ("FFEB9C" if h > 40 else (cor_sem or cor_row))
                _cel(ws, row, 4+i_sem, f"{h}h" if h else "—", bg=cor_h, sz=9)

            _cel(ws, row, 4+n_sem, f"{total}h", bold=True, bg=C["TOTAL"], sz=9)
            _cel(ws, row, 5+n_sem, str(rh.get("cinderelas",0)), bg=cor_row, sz=9)
            _cel(ws, row, 6+n_sem, str(rh.get("plantoes_enf_fds",0)), bg=C["ENF_FDS_H"], sz=9)
            _cel(ws, row, 7+n_sem, str(rh.get("plantoes_ps_fds",0)), bg=C["PS_FDS_H"], sz=9)
            ws.row_dimensions[row].height = 15
            row += 1

    # Rodapé
    row += 1
    nota = f"CH padrão: {config.get('regras_especiais',{}).get('limite_ch',40)}h/sem | Absoluto: {config.get('regras_especiais',{}).get('limite_abs',43)}h | Vermelho > 43h | Amarelo > 40h"
    _cel(ws, row, 1, nota, halign="left", italic=True, bg="F2F2F2", sz=8, border=False)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4+n_sem)

def _cor_semana_aluno(nome, sem_num, resumo_horas, locais_cfg):
    """Tenta retornar cor do local do aluno naquela semana."""
    return None  # simplificado — pode ser expandido


# ── ABA 6: REGRAS E RESTRIÇÕES ────────────────────────────────────────────────
def _aba_regras(wb, titulo, config, dados):
    ws = wb.create_sheet("Regras e Restrições")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 55

    esp = config.get("especialidade","")
    grupo = config.get("grupo","")
    turma = config.get("turma","")
    locais_cfg = config.get("locais",[])
    regras = config.get("regras_especiais",{})

    _header(ws, 1, 1, f"REGRAS E RESTRIÇÕES — {esp} — {grupo} / {turma}", span=3, sz=11)
    for i, h in enumerate(["CATEGORIA","ITEM","DETALHE"],1):
        _cel(ws, 2, i, h, bold=True, bg=C["H2"], fc="FFFFFF", sz=9)

    linhas = [
        ("RODÍZIO", "Descrição", config.get("rodizio_desc","")),
        ("QUINTA-FEIRA", "Regra", regras.get("quinta","Sem tarde — ENAMED")),
        ("TERÇA-FEIRA", "Regra", regras.get("terca","Tarde encurtada")),
        ("CH", "Limite padrão", f"{regras.get('limite_ch',40)}h/semana"),
        ("CH", "Limite absoluto", f"{regras.get('limite_abs',43)}h/semana"),
        ("FDS", "Regras", regras.get("fds","")),
    ]
    for l in locais_cfg:
        bloqs = l.get("bloqueios_tarde",[])
        bloq_str = " | ".join([f"{b['dia']}: {b['tipo']}{' '+b.get('horario','') if b.get('horario') else ''}" for b in bloqs])
        linhas.append((l.get("nome",""), "Manhã", f"{l.get('manha','')} | Mín:{l.get('min_manha',0)} Máx:{l.get('max_manha',0)}"))
        linhas.append((l.get("nome",""), "Tarde", f"{l.get('tarde','')} | Mín:{l.get('min_tarde',0)} Máx:{l.get('max_tarde',0)}"))
        if bloqs: linhas.append((l.get("nome",""), "Bloqueios tarde", bloq_str))
        if l.get("cinderela"): linhas.append((l.get("nome",""), "Cinderela", l.get("cinderela","")))
        if l.get("fds"):
            linhas.append((l.get("nome",""), "FDS — quem faz", l.get("fds_quem","")))
            linhas.append((l.get("nome",""), "FDS — compensação", l.get("fds_comp","")))

    cores_cat = {}; cores_cycle = [C["ENF"],C["PS"],C["AMB"],C["LOC4"],C["LOC5"],"F2F2F2"]
    cor_i = 0
    row = 3
    for cat, item, det in linhas:
        if not det: continue
        if cat not in cores_cat:
            cores_cat[cat] = cores_cycle[cor_i % len(cores_cycle)]; cor_i+=1
        cor = cores_cat[cat]
        _cel(ws, row, 1, cat, bold=True, bg=cor, halign="left", sz=9)
        _cel(ws, row, 2, item, bg=cor, halign="left", sz=9)
        _cel(ws, row, 3, det, bg=cor, halign="left", sz=9, wrap=True)
        ws.row_dimensions[row].height = 15
        row += 1


# ── ABA 7: ESCALA POR SUBGRUPO ────────────────────────────────────────────────
def _aba_escala_subgrupo(wb, titulo, alunos_por_sg, escala_por_aluno, config, semanas, locais_cfg):
    ws = wb.create_sheet("Escala por Subgrupo")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "C6"

    grupo = config.get("grupo","")
    turma = config.get("turma","")
    d_ini = semanas[0][0].strftime("%d/%m/%Y") if semanas else ""
    d_fim = semanas[-1][-1].strftime("%d/%m/%Y") if semanas else ""

    _header(ws, 1, 1, f"ESCALA POR SUBGRUPO — {grupo} / {turma}  |  {d_ini}–{d_fim}  |  rodízio equilibrado",
            span=60, sz=11)
    _header(ws, 2, 1, "M=Manhã · T=Tarde · R=Tarde12-16h · F=FDS manhã · ★=Plantão FDS",
            span=60, bg=C["H2"], sz=9)

    # Montar todas as datas
    todas_datas = []
    for sem in semanas:
        todas_datas.extend(sem)
    n_datas = len(todas_datas)

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 28
    for i in range(n_datas):
        ws.column_dimensions[get_column_letter(3+i)].width = 9

    # Linha 3: Semanas (agrupadas)
    _cel(ws, 3, 1, "SG", bold=True, bg=C["H1"], fc="FFFFFF", sz=9)
    _cel(ws, 3, 2, "Nome", bold=True, bg=C["H1"], fc="FFFFFF", sz=9)
    col = 3
    for i_sem, sem in enumerate(semanas):
        d0 = sem[0]; d1 = sem[-1]
        label = f"S{i_sem+1}|{d0.strftime('%d/%m')}–{d1.strftime('%d/%m')}"
        _header(ws, 3, col, label, span=7, bg=C["H1"], sz=8)
        col += 7

    # Linha 4: Datas
    _cel(ws, 4, 1, "", bg=C["H2"])
    _cel(ws, 4, 2, "", bg=C["H2"])
    for i, dt in enumerate(todas_datas):
        eh_fds = dt.weekday() >= 5
        bg = C["FDS"] if eh_fds else C["H3"]
        _cel(ws, 4, 3+i, dt.strftime("%d/%m"), bold=True, bg=bg, sz=8)

    # Linha 5: Dias da semana
    _cel(ws, 5, 1, "", bg=C["H2"])
    _cel(ws, 5, 2, "", bg=C["H2"])
    for i, dt in enumerate(todas_datas):
        eh_fds = dt.weekday() >= 5
        bg = C["FDS"] if eh_fds else C["H3"]
        _cel(ws, 5, 3+i, DIAS_PT[dt.weekday()], bold=True, bg=bg, sz=8)

    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 16
    ws.row_dimensions[5].height = 14

    row = 6
    for sg_key in sorted(alunos_por_sg.keys(), key=lambda x: int(x[0]) if x[0].isdigit() else 0):
        sg_num = int(sg_key) if sg_key.isdigit() else 1
        cor_sg = _cor_sg(sg_num)
        nomes = alunos_por_sg[sg_key]

        # Linha de separador do SG
        label_sg = f"SG{sg_num} — " + " · ".join([n.split()[0] for n in nomes[:4]])
        _header(ws, row, 1, label_sg, span=2+n_datas, bg=C["H2"], sz=9)
        row += 1

        for nome in nomes:
            _cel(ws, row, 1, f"SG{sg_num}", bold=True, bg=cor_sg, sz=8)
            _cel(ws, row, 2, nome, halign="left", bg=cor_sg, sz=8)

            for i, dt in enumerate(todas_datas):
                data_str = dt.strftime("%d/%m/%Y")
                data_str2 = dt.strftime("%d/%m")
                data_str3 = dt.strftime("%Y-%m-%d")
                aluno_datas = escala_por_aluno.get(nome, {})
                valor = (aluno_datas.get(data_str) or
                         aluno_datas.get(data_str2) or
                         aluno_datas.get(data_str3) or "—")

                # Cor da célula
                if valor == "—":
                    cor_cel = C["FDS_CELL"] if dt.weekday() >= 5 else "FFFFFF"
                else:
                    local_nome = valor.split("(")[0] if "(" in valor else valor
                    cor_cel = _cor_local(local_nome, locais_cfg)
                    if "FDS" in valor or "★" in valor: cor_cel = C["FDS_PLT"]
                    elif "(R)" in valor or "+R" in valor: cor_cel = C["TARDE_R"]

                _cel(ws, row, 3+i, valor, bg=cor_cel, sz=7)

            ws.row_dimensions[row].height = 13
            row += 1


# ── ABA 8: ESCALA INDIVIDUAL ──────────────────────────────────────────────────
def _aba_escala_individual(wb, titulo, alunos_por_sg, escala_por_aluno, config, semanas, locais_cfg):
    ws = wb.create_sheet("Escala Individual")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "D4"

    grupo = config.get("grupo","")
    turma = config.get("turma","")
    d_ini = semanas[0][0].strftime("%d/%m/%Y") if semanas else ""
    d_fim = semanas[-1][-1].strftime("%d/%m/%Y") if semanas else ""

    _header(ws, 1, 1, f"ESCALA INDIVIDUAL — {grupo} / {turma}  |  {d_ini}–{d_fim}  |  rodízio equilibrado",
            span=60, sz=11)
    _header(ws, 2, 1, "M=Manhã · T=Tarde · R=Tarde12-16h · F=FDS manhã · ★=Plantão PS FDS",
            span=60, bg=C["H2"], sz=9)

    # Montar todas as datas
    todas_datas = []
    for sem in semanas:
        todas_datas.extend(sem)
    n_datas = len(todas_datas)

    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 5
    ws.column_dimensions["C"].width = 12
    for i in range(n_datas):
        ws.column_dimensions[get_column_letter(4+i)].width = 9

    # Linha 3: headers
    for ci, h in enumerate(["Nome","SG","RA"],1):
        _cel(ws, 3, ci, h, bold=True, bg=C["H1"], fc="FFFFFF", sz=9)
    for i, dt in enumerate(todas_datas):
        eh_fds = dt.weekday() >= 5
        bg = C["FDS"] if eh_fds else C["H3"]
        _cel(ws, 3, 4+i, dt.strftime("%d/%m"), bold=True, bg=bg, sz=8)
    ws.row_dimensions[3].height = 16

    row = 4
    for sg_key in sorted(alunos_por_sg.keys(), key=lambda x: int(x[0]) if x[0].isdigit() else 0):
        sg_num = int(sg_key) if sg_key.isdigit() else 1
        cor_sg = _cor_sg(sg_num)

        for i_al, nome in enumerate(alunos_por_sg[sg_key]):
            cor_nome = C["ALT1"] if i_al % 2 == 0 else C["ALT2"]
            _cel(ws, row, 1, nome, halign="left", bg=cor_sg, bold=True, sz=9)
            _cel(ws, row, 2, f"SG{sg_num}", bold=True, bg=cor_sg, sz=9)
            _cel(ws, row, 3, "", bg=cor_nome, sz=9)  # RA

            for i, dt in enumerate(todas_datas):
                data_str = dt.strftime("%d/%m/%Y")
                data_str2 = dt.strftime("%d/%m")
                data_str3 = dt.strftime("%Y-%m-%d")
                aluno_datas = escala_por_aluno.get(nome, {})
                valor = (aluno_datas.get(data_str) or
                         aluno_datas.get(data_str2) or
                         aluno_datas.get(data_str3) or "—")

                if valor == "—":
                    cor_cel = C["FDS_CELL"] if dt.weekday() >= 5 else "FFFFFF"
                else:
                    local_nome = valor.split("(")[0] if "(" in valor else valor
                    cor_cel = _cor_local(local_nome, locais_cfg)
                    if "FDS" in valor or "★" in valor: cor_cel = C["FDS_PLT"]
                    elif "(R)" in valor or "+R" in valor: cor_cel = C["TARDE_R"]

                _cel(ws, row, 4+i, valor, bg=cor_cel, sz=7)

            ws.row_dimensions[row].height = 14
            row += 1

# v2 - formato CM5 exato
