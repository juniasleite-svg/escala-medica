"""
Gerador de Excel de Escala Médica — formato CM5
Recebe dados JSON da IA e produz Excel com 9 abas
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
from datetime import datetime

# ── Paleta de cores ──────────────────────────────────────────────────────────
COR = {
    "header_dark":  "1F4E79",
    "header_med":   "2E75B6",
    "header_light": "BDD7EE",
    "enf":          "D6E4F7",  # azul claro
    "ps":           "D6F7D6",  # verde claro
    "amb":          "FAD7A0",  # laranja
    "fds":          "D7BDE2",  # roxo claro
    "fds_ps":       "C39BD3",  # roxo médio (PA★)
    "tarde_r":      "FFF2CC",  # amarelo (tarde reduzida)
    "verde":        "E2EFDA",  # área verde
    "cinderela":    "F4B942",  # laranja escuro
    "feriado":      "FFCCCC",  # vermelho claro
    "sg1": "D6E4F7", "sg2": "D6F7D6", "sg3": "FAD7A0",
    "sg4": "F7D6D6", "sg5": "F2D9F7", "sg6": "D9F7E8",
    "sg7": "FDE8D8", "sg8": "E8F8E8",
    "branco": "FFFFFF", "cinza_claro": "F2F2F2", "cinza": "D9D9D9",
}

def thin_border(color="BFBFBF"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def med_border():
    s = Side(style="medium", color="1F4E79")
    return Border(left=s, right=s, top=s, bottom=s)

def c(ws, row, col, value="", bold=False, bg=None, fc="000000",
      align="center", wrap=False, sz=9, italic=False, border=True):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(bold=bold, color=fc, size=sz, italic=italic)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if border:
        cell.border = thin_border()
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    return cell

def header(ws, row, col, value, span=None, bg=COR["header_dark"], fc="FFFFFF", sz=11):
    cell = c(ws, row, col, value, bold=True, bg=bg, fc=fc, sz=sz)
    if span and span > 1:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+span-1)
    return cell

def cor_local(local_nome, locais_config):
    """Retorna cor baseada no nome do local."""
    cores = [COR["enf"], COR["ps"], COR["amb"], COR["fds"],
             COR["sg4"], COR["sg5"], COR["sg6"], COR["sg7"]]
    nomes = [l.get("nome","") for l in locais_config]
    try:
        idx = nomes.index(local_nome)
        return cores[idx % len(cores)]
    except:
        return COR["cinza_claro"]

def cor_sg(sg_num):
    key = f"sg{sg_num}"
    return COR.get(key, COR["cinza_claro"])


def gerar_excel_completo(dados, config):
    """
    dados: dict com keys: calendario_rodizio, escala_detalhada, resumo_horas, confirmacao, auditoria
    config: dict com especialidade, grupo, turma, data_inicio, num_semanas, locais, alunos_por_sg, ...
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    esp = config.get("especialidade","")
    grupo = config.get("grupo","")
    turma = config.get("turma","")
    titulo = f"{esp} — {config.get('ano_curso','')} — {grupo} / {turma}"

    locais_cfg = config.get("locais", [])
    alunos_por_sg = config.get("alunos_por_sg", {})
    num_semanas = int(config.get("num_semanas", 8))

    _aba_resumo_geral(wb, titulo, config, dados)
    _aba_alunos(wb, titulo, alunos_por_sg, config)
    _aba_calendario(wb, titulo, dados.get("calendario_rodizio", []), config)
    _aba_nominal_detalhada(wb, titulo, dados.get("escala_detalhada", []), config)
    _aba_resumo_horas(wb, titulo, dados.get("resumo_horas", []), config)
    _aba_regras(wb, titulo, config, dados)
    _aba_por_subgrupo(wb, titulo, dados.get("escala_detalhada", []), config)
    _aba_individual(wb, titulo, dados.get("escala_detalhada", []), config)
    _aba_por_servico(wb, titulo, dados.get("escala_detalhada", []), config)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()


def _aba_resumo_geral(wb, titulo, config, dados):
    ws = wb.create_sheet("Resumo Geral")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 55

    header(ws, 1, 1, titulo, span=2, sz=14)
    ws.row_dimensions[1].height = 30

    locais_cfg = config.get("locais", [])
    rodizio_desc = config.get("rodizio_desc", "")
    regras = config.get("regras_especiais", {})

    dados_resumo = [
        ("Período", f"{config.get('data_inicio','')} ({config.get('num_semanas',8)} semanas)"),
        ("Grupo / Turma", f"{config.get('grupo','')} — {config.get('turma','')} — {config.get('ano_curso','')}"),
        ("Total de alunos", str(sum(len(v) for v in config.get("alunos_por_sg",{}).values()))),
        ("Subgrupos", ", ".join([f"SG{k}" for k in sorted(config.get("alunos_por_sg",{}).keys(), key=lambda x: int(x))])),
        ("Locais", " · ".join([l.get("nome","") for l in locais_cfg])),
        ("Rodízio", rodizio_desc or dados.get("confirmacao","")[:100]),
        ("Quinta-feira", regras.get("quinta","ENAMED — sem tarde")),
        ("Terça-feira", regras.get("terca","Aula — tarde 12-16h")),
        ("Limite CH", f"{regras.get('limite_ch',40)}h/sem | Absoluto: {regras.get('limite_abs',43)}h"),
    ]

    row = 3
    for label, val in dados_resumo:
        c(ws, row, 1, label, bold=True, bg=COR["header_light"], align="left", sz=9)
        c(ws, row, 2, val, align="left", sz=9)
        ws.row_dimensions[row].height = 16
        row += 1

    # Legenda de locais com cores
    row += 1
    header(ws, row, 1, "LEGENDA DE LOCAIS", span=2, bg=COR["header_med"])
    row += 1
    cores_loc = [COR["enf"], COR["ps"], COR["amb"], COR["fds"], COR["sg4"], COR["sg5"]]
    for i, local in enumerate(locais_cfg):
        cor = cores_loc[i % len(cores_loc)]
        c(ws, row, 1, local.get("abrev", local.get("nome","")), bold=True, bg=cor, sz=9)
        obs = f"{local.get('nome','')} | M:{local.get('manha','')} T:{local.get('tarde','')} C:{local.get('cinderela','—')}"
        c(ws, row, 2, obs, align="left", bg=cor, sz=9)
        row += 1

    # Legenda de turnos
    row += 1
    header(ws, row, 1, "LEGENDA DE TURNOS", span=2, bg=COR["header_med"])
    row += 1
    turnos = [
        ("M", "Manhã (horário conforme local)", COR["branco"]),
        ("T", "Tarde completa", COR["branco"]),
        ("R", "Tarde reduzida 12-16h (fundo amarelo)", COR["tarde_r"]),
        ("F", "FDS manhã Enf/PS", COR["fds"]),
        ("PA★", "Plantão PA no FDS — alunos do Ambulatório (fundo roxo)", COR["fds_ps"]),
        ("AV", "Área Verde — tarde livre", COR["verde"]),
        ("C", "Cinderela", COR["cinderela"]),
        ("—", "Sem atividade (FDS sem plantão)", COR["cinza_claro"]),
    ]
    for abrev, desc, cor_bg in turnos:
        c(ws, row, 1, abrev, bold=True, bg=cor_bg, sz=9)
        c(ws, row, 2, desc, align="left", bg=cor_bg, sz=9)
        row += 1


def _aba_alunos(wb, titulo, alunos_por_sg, config):
    ws = wb.create_sheet("Alunos")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 38
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 8
    ws.column_dimensions["F"].width = 10

    header(ws, 1, 1, f"LISTA DE ALUNOS — {titulo}", span=6, sz=12)
    ws.row_dimensions[1].height = 26
    for i, h in enumerate(["Nº", "SG", "Nome Completo", "RA", "Turma", "Par"], 1):
        c(ws, 2, i, h, bold=True, bg=COR["header_med"], fc="FFFFFF")

    row = 3
    num = 1
    for sg_key in sorted(alunos_por_sg.keys(), key=lambda x: int(x)):
        sg_num = int(sg_key)
        cor = cor_sg(sg_num)
        for nome in alunos_por_sg[sg_key]:
            c(ws, row, 1, num, bg=cor)
            c(ws, row, 2, f"SG{sg_num}", bold=True, bg=cor)
            c(ws, row, 3, nome, align="left", bg=cor)
            c(ws, row, 4, "", bg=cor)  # RA
            c(ws, row, 5, config.get("turma",""), bg=cor)
            c(ws, row, 6, "", bg=cor)  # Par
            num += 1
            row += 1

    # Pares
    row += 1
    header(ws, row, 1, "PARES", span=6, bg=COR["header_med"])
    row += 1
    for par in config.get("pares", []):
        c(ws, row, 1, par.get("nome",""), bold=True, bg=COR["cinza_claro"], span=None)
        c(ws, row, 2, par.get("sgs",""), bg=COR["cinza_claro"])
        c(ws, row, 3, par.get("rodizio",""), align="left", bg=COR["cinza_claro"])
        row += 1


def _aba_calendario(wb, titulo, calendario, config):
    ws = wb.create_sheet("Calendário de Rodízio")
    ws.sheet_view.showGridLines = False

    locais_cfg = config.get("locais", [])
    alunos_por_sg = config.get("alunos_por_sg", {})
    sgs = sorted(alunos_por_sg.keys(), key=lambda x: int(x))

    header(ws, 1, 1, f"CALENDÁRIO DE RODÍZIO — {titulo}", span=2+len(sgs), sz=12)
    ws.row_dimensions[1].height = 26
    c(ws, 2, 1, "Par / SG", bold=True, bg=COR["header_med"], fc="FFFFFF")
    c(ws, 2, 2, "Período", bold=True, bg=COR["header_med"], fc="FFFFFF")
    for i, sg in enumerate(sgs):
        n_al = len(alunos_por_sg.get(sg, []))
        c(ws, 2, 3+i, f"SG{sg}({n_al})", bold=True, bg=COR["header_med"], fc="FFFFFF")

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 22
    for i in range(len(sgs)):
        ws.column_dimensions[get_column_letter(3+i)].width = 20

    cores_loc = [COR["enf"], COR["ps"], COR["amb"], COR["fds"], COR["sg4"], COR["sg5"]]
    loc_nomes = [l.get("nome","") for l in locais_cfg]
    loc_cores = {n: cores_loc[i % len(cores_loc)] for i, n in enumerate(loc_nomes)}

    for row_i, sem in enumerate(calendario):
        row = row_i + 3
        c(ws, row, 1, f"Sem {sem.get('semana','')}", bold=True, bg=COR["cinza_claro"])
        c(ws, row, 2, sem.get("periodo",""), bg=COR["cinza_claro"])
        aloc = sem.get("alocacao", {})
        for i, sg in enumerate(sgs):
            local = aloc.get(f"SG{sg}", aloc.get(sg, ""))
            cor = loc_cores.get(local, COR["cinza_claro"])
            c(ws, row, 3+i, local, bg=cor)

    # Blocos
    row = len(calendario) + 4
    header(ws, row, 1, "BLOCOS DE RODÍZIO", span=2+len(sgs), bg=COR["header_med"])
    for bloco in config.get("blocos", []):
        row += 1
        c(ws, row, 1, bloco.get("nome",""), bold=True, bg=COR["cinza_claro"])
        c(ws, row, 2, bloco.get("periodo",""), bg=COR["cinza_claro"])
        c(ws, row, 3, bloco.get("descricao",""), align="left", bg=COR["cinza_claro"])
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=2+len(sgs))


def _aba_nominal_detalhada(wb, titulo, escala, config):
    ws = wb.create_sheet("Escala Nominal Detalhada")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "E3"

    regras = config.get("regras_especiais", {})
    subtitulo = f"M={regras.get('manha_horario','07-13h')}  ·  FDS={regras.get('fds_desc','')}  ·  Terça={regras.get('terca','')}  ·  Qui={regras.get('quinta','')}"

    header(ws, 1, 1, f"ESCALA NOMINAL DETALHADA — {titulo}", span=10, sz=12)
    c(ws, 1, 1, f"ESCALA NOMINAL DETALHADA — {titulo}\n{subtitulo}", bold=True,
      bg=COR["header_dark"], fc="FFFFFF", wrap=True, sz=10)
    ws.merge_cells("A1:J1")
    ws.row_dimensions[1].height = 36

    headers = ["Sem", "Data", "Dia", "Local", "Turno", "Horário", "h", "SG", "Nome", "RA"]
    for i, h in enumerate(headers, 1):
        c(ws, 2, i, h, bold=True, bg=COR["header_med"], fc="FFFFFF")

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 6
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 5
    ws.column_dimensions["H"].width = 6
    ws.column_dimensions["I"].width = 36
    ws.column_dimensions["J"].width = 14

    locais_cfg = config.get("locais", [])
    cores_loc_list = [COR["enf"], COR["ps"], COR["amb"], COR["fds"], COR["sg4"], COR["sg5"]]
    loc_cores = {l.get("nome",""): cores_loc_list[i % len(cores_loc_list)] for i, l in enumerate(locais_cfg)}

    for row_i, entry in enumerate(escala):
        row = row_i + 3
        local = entry.get("local","")
        turno = entry.get("turno","")
        cor = loc_cores.get(local, COR["cinza_claro"])
        if "FDS" in turno or "★" in turno or "PA" in turno:
            cor = COR["fds_ps"]
        elif "Reduz" in turno or turno == "R":
            cor = COR["tarde_r"]
        elif "Verde" in turno or turno == "AV":
            cor = COR["verde"]
        elif "Cinderela" in turno or turno == "C":
            cor = COR["cinderela"]

        vals = [
            entry.get("semana",""), entry.get("data",""), entry.get("dia",""),
            local, turno, entry.get("horario",""), entry.get("horas",""),
            entry.get("sg",""), entry.get("nome","") if isinstance(entry.get("nome"), str) else " | ".join(entry.get("alunos",[])),
            entry.get("ra","")
        ]
        for col_i, v in enumerate(vals, 1):
            c(ws, row, col_i, v, bg=cor, align="left" if col_i in [4,9] else "center")
        ws.row_dimensions[row].height = 14


def _aba_resumo_horas(wb, titulo, resumo, config):
    ws = wb.create_sheet("Resumo de Horas")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "D3"

    num_semanas = int(config.get("num_semanas", 8))
    header(ws, 1, 1, f"RESUMO DE HORAS — {titulo}", span=6+num_semanas, sz=12)
    ws.row_dimensions[1].height = 26

    sem_headers = [f"S{i+1}" for i in range(num_semanas)]
    headers = ["SG","Nome","RA"] + sem_headers + ["TOTAL","Cind","Enf FDS","PS FDS"]
    for i, h in enumerate(headers, 1):
        bg = COR["header_dark"] if h in ["TOTAL","SG"] else COR["header_med"]
        c(ws, 2, i, h, bold=True, bg=bg, fc="FFFFFF", sz=9)

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 14
    for i in range(num_semanas):
        ws.column_dimensions[get_column_letter(4+i)].width = 8
    ws.column_dimensions[get_column_letter(4+num_semanas)].width = 10
    ws.column_dimensions[get_column_letter(5+num_semanas)].width = 8
    ws.column_dimensions[get_column_letter(6+num_semanas)].width = 10
    ws.column_dimensions[get_column_letter(7+num_semanas)].width = 10

    for row_i, aluno in enumerate(resumo):
        row = row_i + 3
        sg_num = int(aluno.get("sg", 1))
        cor = cor_sg(sg_num)
        semanas_h = aluno.get("semanas", [])
        total = aluno.get("total_horas", sum(semanas_h) if semanas_h else 0)

        c(ws, row, 1, f"SG{sg_num}", bold=True, bg=cor)
        c(ws, row, 2, aluno.get("nome",""), align="left", bg=cor)
        c(ws, row, 3, str(aluno.get("ra","")), bg=cor)
        for i, h in enumerate(semanas_h):
            cor_h = "FFC7CE" if h > 43 else ("FFEB9C" if h > 40 else cor)
            c(ws, row, 4+i, f"{h}h", bg=cor_h)
        c(ws, row, 4+len(semanas_h), f"{total}h", bold=True, bg=cor)
        c(ws, row, 5+len(semanas_h), aluno.get("cinderelas", 0), bg=cor)
        c(ws, row, 6+len(semanas_h), aluno.get("plantoes_enf_fds", aluno.get("plantoes_fds",0)), bg=cor)
        c(ws, row, 7+len(semanas_h), aluno.get("plantoes_ps_fds", 0), bg=cor)
        ws.row_dimensions[row].height = 15

    # Rodapé
    row = len(resumo) + 3
    c(ws, row, 1, "Nota:", bold=True, align="left", bg=COR["cinza_claro"])
    nota = config.get("regras_especiais",{}).get("nota_ch","40h padrão · >43h = crítico (vermelho) · >40h = atenção (amarelo)")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4+num_semanas)
    c(ws, row, 1, nota, align="left", italic=True, bg=COR["cinza_claro"], sz=8)


def _aba_regras(wb, titulo, config, dados):
    ws = wb.create_sheet("Regras e Restrições")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 55

    header(ws, 1, 1, f"REGRAS E RESTRIÇÕES — {titulo}", span=3, sz=12)
    ws.row_dimensions[1].height = 26
    for i, h in enumerate(["CATEGORIA","ITEM","DETALHE"], 1):
        c(ws, 2, i, h, bold=True, bg=COR["header_med"], fc="FFFFFF")

    regras = config.get("regras_especiais", {})
    locais_cfg = config.get("locais", [])

    linhas = [
        ("RODÍZIO", "Descrição", config.get("rodizio_desc","")),
        ("QUINTA-FEIRA", "ENAMED", regras.get("quinta","Sem tarde")),
        ("TERÇA-FEIRA", "Aula", regras.get("terca","Tarde 12-16h")),
        ("CH", "Limite padrão", f"{regras.get('limite_ch',40)}h/semana"),
        ("CH", "Limite absoluto", f"{regras.get('limite_abs',43)}h/semana"),
        ("FDS", "Regras", regras.get("fds","")),
        ("ÁREA VERDE", "Direito", "1 período livre/semana/aluno"),
        ("ÁREA VERDE", "Quando", "Preferencialmente à tarde"),
    ]
    for local in locais_cfg:
        linhas.append((local.get("nome",""), "Manhã", local.get("manha","")))
        linhas.append((local.get("nome",""), "Tarde", local.get("tarde","")))
        if local.get("cinderela"):
            linhas.append((local.get("nome",""), "Cinderela", local.get("cinderela","")))
        if local.get("obs"):
            linhas.append((local.get("nome",""), "Obs", local.get("obs","")))

    cores_cat = {}
    cor_cycle = [COR["enf"], COR["ps"], COR["amb"], COR["fds"], COR["cinza_claro"]]
    cor_idx = 0
    for row_i, (cat, item, det) in enumerate(linhas):
        row = row_i + 3
        if cat not in cores_cat:
            cores_cat[cat] = cor_cycle[cor_idx % len(cor_cycle)]
            cor_idx += 1
        cor = cores_cat[cat]
        c(ws, row, 1, cat, bold=True, bg=cor, align="left")
        c(ws, row, 2, item, bg=cor, align="left")
        c(ws, row, 3, det, bg=cor, align="left", wrap=True)
        ws.row_dimensions[row].height = 16

    # Auditoria
    row = len(linhas) + 4
    header(ws, row, 1, "AUDITORIA", span=3, bg=COR["header_dark"])
    audit = dados.get("auditoria", {})
    row += 1
    status = "✅ APROVADA" if audit.get("aprovado") else "❌ COM ERROS"
    c(ws, row, 1, "Status", bold=True, bg=COR["cinza_claro"])
    c(ws, row, 2, status, bold=True, bg="C6EFCE" if audit.get("aprovado") else "FFC7CE", span=None)
    for err in audit.get("erros", []):
        row += 1
        c(ws, row, 1, "ERRO", bold=True, bg="FFC7CE")
        c(ws, row, 2, err, align="left", bg="FFC7CE")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
    for av in audit.get("avisos", []):
        row += 1
        c(ws, row, 1, "AVISO", bold=True, bg="FFEB9C")
        c(ws, row, 2, av, align="left", bg="FFEB9C")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)


def _aba_por_subgrupo(wb, titulo, escala, config):
    ws = wb.create_sheet("Escala por Subgrupo")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 36

    header(ws, 1, 1, f"ESCALA POR SUBGRUPO — {titulo}", span=10, sz=12)
    ws.row_dimensions[1].height = 26

    alunos_por_sg = config.get("alunos_por_sg", {})
    locais_cfg = config.get("locais", [])
    cores_loc_list = [COR["enf"], COR["ps"], COR["amb"], COR["fds"], COR["sg4"], COR["sg5"]]
    loc_cores = {l.get("nome",""): cores_loc_list[i % len(cores_loc_list)] for i, l in enumerate(locais_cfg)}

    # Agrupar escala por sg+aluno+semana
    from collections import defaultdict
    por_aluno = defaultdict(lambda: defaultdict(list))
    for entry in escala:
        sg = str(entry.get("sg",""))
        nome = entry.get("nome","") if isinstance(entry.get("nome"), str) else ""
        sem = entry.get("semana", 0)
        por_aluno[sg][nome].append(entry)

    row = 3
    for sg_key in sorted(alunos_por_sg.keys(), key=lambda x: int(x)):
        sg_num = int(sg_key)
        cor = cor_sg(sg_num)
        nomes_sg = alunos_por_sg[sg_key]

        # Cabeçalho do SG
        c(ws, row, 1, f"SG{sg_num}", bold=True, bg=COR["header_med"], fc="FFFFFF", sz=10)
        label = " · ".join(nomes_sg[:4]) + ("..." if len(nomes_sg)>4 else "")
        c(ws, row, 2, f"SG{sg_num} — {label}", bold=True, bg=COR["header_med"], fc="FFFFFF", sz=10, align="left")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=10)
        row += 1

        for nome in nomes_sg:
            c(ws, row, 1, "", bg=cor)
            c(ws, row, 2, nome, bold=False, bg=cor, align="left")
            col = 3
            # Agrupar por semana
            entradas_aluno = [e for e in escala if
                              str(e.get("sg",""))==sg_key and
                              (e.get("nome","")==nome or nome in e.get("alunos",[]))]
            por_sem = defaultdict(list)
            for e in entradas_aluno:
                por_sem[e.get("semana",0)].append(e)

            for sem_num in range(1, int(config.get("num_semanas",8))+1):
                entries = por_sem.get(sem_num, [])
                if entries:
                    # Resumir: local + turnos
                    local = entries[0].get("local","")
                    turnos = "+".join(set(e.get("turno","")[0] for e in entries if e.get("turno")))
                    txt = f"{local[:3]}({turnos})" if local else "—"
                    cor_cell = loc_cores.get(local, cor)
                else:
                    txt = "—"
                    cor_cell = cor
                c(ws, row, col, txt, bg=cor_cell, sz=8)
                if col == 3:
                    ws.column_dimensions[get_column_letter(col)].width = 14
                col += 1
            row += 1
        row += 1


def _aba_individual(wb, titulo, escala, config):
    ws = wb.create_sheet("Escala Individual")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 14

    header(ws, 1, 1, f"ESCALA INDIVIDUAL — {titulo}", span=10, sz=12)
    ws.row_dimensions[1].height = 26

    alunos_por_sg = config.get("alunos_por_sg", {})
    locais_cfg = config.get("locais", [])
    cores_loc_list = [COR["enf"], COR["ps"], COR["amb"], COR["fds"], COR["sg4"], COR["sg5"]]
    loc_cores = {l.get("nome",""): cores_loc_list[i % len(cores_loc_list)] for i, l in enumerate(locais_cfg)}

    from collections import defaultdict
    por_aluno_data = defaultdict(dict)
    for entry in escala:
        nome = entry.get("nome","") if isinstance(entry.get("nome"), str) else ""
        for n in (entry.get("alunos", [nome]) if not nome else [nome]):
            data = entry.get("data","")
            if data:
                if data not in por_aluno_data[n]:
                    por_aluno_data[n][data] = []
                por_aluno_data[n][data].append(entry)

    row = 3
    for sg_key in sorted(alunos_por_sg.keys(), key=lambda x: int(x)):
        sg_num = int(sg_key)
        for nome in alunos_por_sg[sg_key]:
            cor = cor_sg(sg_num)
            c(ws, row, 1, nome, bold=True, bg=cor, align="left", sz=10)
            c(ws, row, 2, f"SG{sg_num}", bold=True, bg=cor)
            c(ws, row, 3, "", bg=cor)  # RA
            row += 1
            dados_aluno = por_aluno_data.get(nome, {})
            for data in sorted(dados_aluno.keys()):
                entries = dados_aluno[data]
                for e in entries:
                    local = e.get("local","")
                    cor_cell = loc_cores.get(local, cor)
                    turno = e.get("turno","")
                    if "FDS" in turno or "★" in turno: cor_cell = COR["fds_ps"]
                    elif "Reduz" in turno or turno=="R": cor_cell = COR["tarde_r"]
                    txt = f"{local} {e.get('horario','')} ({e.get('horas','')}h)"
                    c(ws, row, 1, txt, align="left", bg=cor_cell, sz=8)
                    c(ws, row, 2, data, bg=cor_cell, sz=8)
                    c(ws, row, 3, e.get("dia",""), bg=cor_cell, sz=8)
                    row += 1
            row += 1


def _aba_por_servico(wb, titulo, escala, config):
    ws = wb.create_sheet("Escala por Serviço")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 40

    header(ws, 1, 1, f"ESCALA POR SERVIÇO — {titulo}", span=6, sz=12)
    ws.row_dimensions[1].height = 26
    for i, h in enumerate(["Local/Serviço","Data","Dia","Turno","Horário","Alunos"], 1):
        c(ws, 2, i, h, bold=True, bg=COR["header_med"], fc="FFFFFF")

    locais_cfg = config.get("locais", [])
    cores_loc_list = [COR["enf"], COR["ps"], COR["amb"], COR["fds"], COR["sg4"], COR["sg5"]]
    loc_cores = {l.get("nome",""): cores_loc_list[i % len(cores_loc_list)] for i, l in enumerate(locais_cfg)}

    # Ordenar por local, depois data
    from collections import defaultdict
    por_local = defaultdict(list)
    for e in escala:
        por_local[e.get("local","")].append(e)

    row = 3
    for local in sorted(por_local.keys()):
        cor = loc_cores.get(local, COR["cinza_claro"])
        entries = sorted(por_local[local], key=lambda x: (x.get("data",""), x.get("turno","")))
        for e in entries:
            alunos = e.get("alunos", [])
            if isinstance(e.get("nome",""), str) and e.get("nome"):
                alunos = [e["nome"]]
            txt_alunos = " | ".join(alunos) if alunos else ""
            vals = [local, e.get("data",""), e.get("dia",""), e.get("turno",""), e.get("horario",""), txt_alunos]
            for col_i, v in enumerate(vals, 1):
                c(ws, row, col_i, v, bg=cor, align="left" if col_i in [1,6] else "center", sz=8)
            ws.row_dimensions[row].height = 13
            row += 1
        row += 1  # espaço entre locais

