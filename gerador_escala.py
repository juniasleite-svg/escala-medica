import pandas as pd
from datetime import date, timedelta
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Cores por local (até 6 locais)
CORES_LOCAIS = [
    "AED6F1",  # azul claro
    "A9DFBF",  # verde claro
    "FAD7A0",  # laranja claro
    "D7BDE2",  # lilás
    "F9E79F",  # amarelo
    "FADBD8",  # rosa
]

COR_CINDERELA = "F4B942"
COR_FDS = "C5CAD6"
COR_BLOQUEIO = "FFCCCC"
COR_HEADER = "1F4E79"
COR_HEADER2 = "2E75B6"


def thin_border():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)


def cell_style(ws, row, col, value, bold=False, bg=None, font_color="000000", align="center", wrap=False, font_size=9):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(bold=bold, color=font_color, size=font_size)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    c.border = thin_border()
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    return c


def gerar_datas(data_inicio_str, num_semanas):
    """Retorna lista de semanas, cada uma com lista de datas (seg a dom)."""
    d0 = date.fromisoformat(data_inicio_str)
    semanas = []
    for s in range(num_semanas):
        inicio = d0 + timedelta(weeks=s)
        semana = [inicio + timedelta(days=d) for d in range(7)]
        semanas.append(semana)
    return semanas


def montar_rodizio(alunos, locais, num_semanas, num_sg):
    """
    Monta o calendário de rodízio: qual SG fica em qual local em cada semana.
    Garante que todo SG passa por todo local.
    """
    n_locais = len(locais)
    rodizio = {}  # semana -> {sg: local_idx}

    for sem in range(num_semanas):
        rodizio[sem] = {}
        for sg in range(1, num_sg + 1):
            local_idx = (sg - 1 + sem) % n_locais
            rodizio[sem][sg] = local_idx

    return rodizio


def gerar_escala(config):
    especialidade = config["especialidade"]
    grupo = config["grupo"]
    turma = config["turma"]
    data_inicio = config["data_inicio"]
    num_semanas = config["num_semanas"]
    num_sg = config["num_sg"]
    locais = config["locais"]
    alunos = config["alunos"]
    bloqueios = config["bloqueios"]

    semanas = gerar_datas(data_inicio, num_semanas)
    rodizio = montar_rodizio(alunos, locais, num_semanas, num_sg)

    # Agrupar alunos por SG
    alunos_por_sg = {}
    for a in alunos:
        sg = int(a["Sub Grupo"])
        alunos_por_sg.setdefault(sg, []).append(a)

    # ── Gerar Excel ──────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _aba_subgrupos(wb, alunos_por_sg, num_sg, locais)
    _aba_calendario(wb, semanas, rodizio, locais, num_sg, bloqueios)
    _aba_nominal(wb, semanas, rodizio, locais, alunos_por_sg, num_sg, bloqueios)
    _aba_resumo(wb, semanas, alunos_por_sg, num_sg, locais, rodizio, bloqueios)

    xlsx_io = io.BytesIO()
    wb.save(xlsx_io)
    xlsx_io.seek(0)

    # ── Gerar PPTX simples ───────────────────────────────────────────────────
    pptx_io = _gerar_pptx(especialidade, grupo, turma, semanas, rodizio, locais, alunos_por_sg, num_sg)

    return {"xlsx": xlsx_io.read(), "pptx": pptx_io.read()}


# ── ABA SUBGRUPOS ────────────────────────────────────────────────────────────
def _aba_subgrupos(wb, alunos_por_sg, num_sg, locais):
    ws = wb.create_sheet("Subgrupos")
    ws.sheet_view.showGridLines = False

    cell_style(ws, 1, 1, "SUBGRUPOS", bold=True, bg=COR_HEADER, font_color="FFFFFF", font_size=12)
    ws.merge_cells("A1:D1")
    ws.row_dimensions[1].height = 30

    headers = ["Sub Grupo", "Nome Completo", "RA", "Cor"]
    for c, h in enumerate(headers, 1):
        cell_style(ws, 2, c, h, bold=True, bg=COR_HEADER2, font_color="FFFFFF")

    row = 3
    cores_sg = ["D6E4F7", "D6F7D6", "F7F0D6", "F7D6D6", "F2D9F7", "D9F7E8", "FDE8D8", "E8F8E8"]
    for sg in range(1, num_sg + 1):
        cor = cores_sg[(sg - 1) % len(cores_sg)]
        for a in alunos_por_sg.get(sg, []):
            cell_style(ws, row, 1, f"SG{sg}", bg=cor)
            cell_style(ws, row, 2, a["Nome Completo"], bg=cor, align="left")
            cell_style(ws, row, 3, str(a["RA"]), bg=cor)
            cell_style(ws, row, 4, "", bg=cor)
            row += 1

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 10


# ── ABA CALENDÁRIO DE RODÍZIO ────────────────────────────────────────────────
def _aba_calendario(wb, semanas, rodizio, locais, num_sg, bloqueios):
    ws = wb.create_sheet("Calendário de Rodízio")
    ws.sheet_view.showGridLines = False
    dias_sem = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]

    cell_style(ws, 1, 1, "CALENDÁRIO DE RODÍZIO", bold=True, bg=COR_HEADER, font_color="FFFFFF", font_size=12)
    ws.merge_cells(f"A1:{get_column_letter(1 + num_sg)}1")
    ws.row_dimensions[1].height = 28

    headers = ["Semana / Período"] + [f"SG{i}" for i in range(1, num_sg + 1)]
    for c, h in enumerate(headers, 1):
        cell_style(ws, 2, c, h, bold=True, bg=COR_HEADER2, font_color="FFFFFF")

    cores_sg = ["D6E4F7", "D6F7D6", "F7F0D6", "F7D6D6", "F2D9F7", "D9F7E8", "FDE8D8", "E8F8E8"]
    cores_locais_map = {i: CORES_LOCAIS[i % len(CORES_LOCAIS)] for i in range(len(locais))}

    for sem_idx, semana in enumerate(semanas):
        row = sem_idx + 3
        inicio = semana[0].strftime("%d/%m")
        fim = semana[-1].strftime("%d/%m")
        cell_style(ws, row, 1, f"Sem {sem_idx+1} | {inicio}–{fim}", bold=True, bg="F2F2F2")
        for sg in range(1, num_sg + 1):
            local_idx = rodizio[sem_idx].get(sg, 0)
            nome_local = locais[local_idx]["nome"] if local_idx < len(locais) else "?"
            cor = cores_locais_map.get(local_idx, "FFFFFF")
            cell_style(ws, row, sg + 1, nome_local, bg=cor)

    ws.column_dimensions["A"].width = 22
    for i in range(1, num_sg + 1):
        ws.column_dimensions[get_column_letter(i + 1)].width = 20


# ── ABA NOMINAL DETALHADA ─────────────────────────────────────────────────────
def _aba_nominal(wb, semanas, rodizio, locais, alunos_por_sg, num_sg, bloqueios):
    ws = wb.create_sheet("Escala Nominal Detalhada")
    ws.sheet_view.showGridLines = False
    dias_sem = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]

    # Cabeçalho
    cell_style(ws, 1, 1, "ESCALA NOMINAL DETALHADA", bold=True, bg=COR_HEADER, font_color="FFFFFF", font_size=12)
    total_cols = 2 + len(semanas) * 7
    ws.merge_cells(f"A1:{get_column_letter(total_cols)}1")
    ws.row_dimensions[1].height = 28

    # Linha 2: SG | Nome | Semanas
    cell_style(ws, 2, 1, "SG", bold=True, bg=COR_HEADER2, font_color="FFFFFF")
    cell_style(ws, 2, 2, "Nome", bold=True, bg=COR_HEADER2, font_color="FFFFFF")

    col = 3
    for sem_idx, semana in enumerate(semanas):
        inicio = semana[0].strftime("%d/%m")
        fim = semana[-1].strftime("%d/%m")
        cell_style(ws, 2, col, f"Sem {sem_idx+1} | {inicio}–{fim}", bold=True, bg=COR_HEADER2, font_color="FFFFFF")
        ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col+6)
        for d, dia in enumerate(semana):
            cell_style(ws, 3, col + d, f"{dias_sem[d]}\n{dia.strftime('%d/%m')}", bold=True, bg="E8EDF5", wrap=True)
        col += 7

    ws.row_dimensions[3].height = 30
    ws.freeze_panes = "C4"

    cores_sg = ["D6E4F7", "D6F7D6", "F7F0D6", "F7D6D6", "F2D9F7", "D9F7E8", "FDE8D8", "E8F8E8"]

    row = 4
    for sg in range(1, num_sg + 1):
        cor_sg = cores_sg[(sg - 1) % len(cores_sg)]
        for aluno in alunos_por_sg.get(sg, []):
            cell_style(ws, row, 1, f"SG{sg}", bg=cor_sg, bold=True)
            cell_style(ws, row, 2, aluno["Nome Completo"], bg=cor_sg, align="left")

            col = 3
            for sem_idx, semana in enumerate(semanas):
                local_idx = rodizio[sem_idx].get(sg, 0)
                local = locais[local_idx] if local_idx < len(locais) else {}
                nome_local = local.get("nome", "?")
                cor_local = CORES_LOCAIS[local_idx % len(CORES_LOCAIS)]

                for d, dia in enumerate(semana):
                    dia_sem = dia.weekday()  # 0=seg, 1=ter, ..., 6=dom
                    eh_fds = dia_sem >= 5

                    if eh_fds and not local.get("fds"):
                        cell_style(ws, row, col + d, "—", bg=COR_FDS)
                    elif bloqueios.get("quinta_tarde") and dia_sem == 3:
                        cell_style(ws, row, col + d, f"{nome_local}/M", bg=cor_local)
                    elif bloqueios.get("terca_parcial") and dia_sem == 1:
                        cell_style(ws, row, col + d, f"{nome_local}/M+T", bg=cor_local)
                    else:
                        turnos = []
                        if local.get("turno_m"): turnos.append("M")
                        if local.get("turno_t"): turnos.append("T")
                        cell_style(ws, row, col + d, f"{nome_local}/{'+'.join(turnos)}", bg=cor_local)
                col += 7
            row += 1

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 36
    for c in range(3, total_cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 10
    ws.row_dimensions[1].height = 28


# ── ABA RESUMO DE HORAS ───────────────────────────────────────────────────────
def _aba_resumo(wb, semanas, alunos_por_sg, num_sg, locais, rodizio, bloqueios):
    ws = wb.create_sheet("Resumo de Horas")
    ws.sheet_view.showGridLines = False

    cell_style(ws, 1, 1, "RESUMO DE HORAS POR ALUNO", bold=True, bg=COR_HEADER, font_color="FFFFFF", font_size=12)
    ws.merge_cells(f"A1:{get_column_letter(4 + len(semanas))}1")
    ws.row_dimensions[1].height = 28

    headers = ["SG", "Nome", "RA", "Total Horas", "Total Cinderelas"] + \
              [f"Sem {s+1}" for s in range(len(semanas))]
    for c, h in enumerate(headers, 1):
        cell_style(ws, 2, c, h, bold=True, bg=COR_HEADER2, font_color="FFFFFF")

    cores_sg = ["D6E4F7", "D6F7D6", "F7F0D6", "F7D6D6", "F2D9F7", "D9F7E8", "FDE8D8", "E8F8E8"]
    row = 3
    for sg in range(1, num_sg + 1):
        cor_sg = cores_sg[(sg - 1) % len(cores_sg)]
        for aluno in alunos_por_sg.get(sg, []):
            horas_semanas = []
            total_cind = 0
            for sem_idx, semana in enumerate(semanas):
                local_idx = rodizio[sem_idx].get(sg, 0)
                local = locais[local_idx] if local_idx < len(locais) else {}
                horas = 0
                cind = 0
                for d, dia in enumerate(semana):
                    dia_sem = dia.weekday()
                    eh_fds = dia_sem >= 5
                    if eh_fds and not local.get("fds"):
                        continue
                    if local.get("turno_m"): horas += 6
                    if bloqueios.get("quinta_tarde") and dia_sem == 3:
                        pass  # sem tarde na quinta
                    elif bloqueios.get("terca_parcial") and dia_sem == 1:
                        horas += 3  # tarde parcial
                    elif local.get("turno_t"):
                        horas += 6
                    if local.get("turno_c") and (eh_fds or dia_sem in [4, 5, 6]):
                        horas += 5
                        cind += 1
                horas_semanas.append(horas)
                total_cind += cind

            cell_style(ws, row, 1, f"SG{sg}", bg=cor_sg, bold=True)
            cell_style(ws, row, 2, aluno["Nome Completo"], bg=cor_sg, align="left")
            cell_style(ws, row, 3, str(aluno["RA"]), bg=cor_sg)
            cell_style(ws, row, 4, sum(horas_semanas), bg=cor_sg, bold=True)
            cell_style(ws, row, 5, total_cind, bg=cor_sg)
            for s, h in enumerate(horas_semanas):
                cell_style(ws, row, 6 + s, h, bg=cor_sg)
            row += 1

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 16
    for s in range(len(semanas)):
        ws.column_dimensions[get_column_letter(6 + s)].width = 10


# ── PPTX SIMPLES ─────────────────────────────────────────────────────────────
def _gerar_pptx(especialidade, grupo, turma, semanas, rodizio, locais, alunos_por_sg, num_sg):
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt, Emu
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN
        import io

        prs = Presentation()
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)

        blank = prs.slide_layouts[6]

        def add_textbox(slide, text, left, top, width, height, bold=False, size=18,
                        color=(0,0,0), bg=None, align=PP_ALIGN.CENTER):
            txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
            tf = txBox.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = align
            run = p.add_run()
            run.text = text
            run.font.bold = bold
            run.font.size = Pt(size)
            run.font.color.rgb = RGBColor(*color)
            if bg:
                fill = txBox.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(*bg)
            return txBox

        # Slide 1: Capa
        slide = prs.slides.add_slide(blank)
        bg = slide.background
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(31, 78, 121)
        add_textbox(slide, f"Escala de Internato\n{especialidade}", 1, 2, 11, 1.5,
                    bold=True, size=32, color=(212, 175, 55))
        add_textbox(slide, f"{grupo} | {turma}", 1, 4, 11, 1,
                    size=20, color=(255,255,255))
        add_textbox(slide, f"{semanas[0][0].strftime('%d/%m/%Y')} a {semanas[-1][-1].strftime('%d/%m/%Y')}",
                    1, 5, 11, 0.8, size=16, color=(200,200,200))

        # Slide 2: Subgrupos
        slide = prs.slides.add_slide(blank)
        add_textbox(slide, "SUBGRUPOS", 0.3, 0.2, 12, 0.6, bold=True, size=22,
                    color=(255,255,255), bg=(31,78,121))
        cores_sg = [(214,228,247),(214,247,214),(247,240,214),(247,214,214),(242,217,247),(217,247,232)]
        col_w = 13.33 / min(num_sg, 4)
        for sg in range(1, num_sg + 1):
            col_idx = (sg - 1) % 4
            row_idx = (sg - 1) // 4
            cor = cores_sg[(sg-1) % len(cores_sg)]
            left = 0.1 + col_idx * col_w
            top = 1.0 + row_idx * 3.0
            nomes = "\n".join([a["Nome Completo"] for a in alunos_por_sg.get(sg, [])])
            add_textbox(slide, f"SG{sg}\n{nomes}", left, top, col_w - 0.2, 2.8,
                        bold=False, size=10, bg=cor, align=PP_ALIGN.LEFT)

        # Slides por semana
        cores_loc = [(174,214,241),(169,223,191),(250,215,160),(215,189,226),(249,231,159),(250,219,216)]
        for sem_idx, semana in enumerate(semanas):
            slide = prs.slides.add_slide(blank)
            inicio = semana[0].strftime("%d/%m")
            fim = semana[-1].strftime("%d/%m")
            add_textbox(slide, f"Semana {sem_idx+1} | {inicio} – {fim}", 0.3, 0.1, 12, 0.55,
                        bold=True, size=18, color=(255,255,255), bg=(31,78,121))

            n_locais = len(locais)
            col_w2 = 13.0 / n_locais
            for loc_idx, local in enumerate(locais):
                cor = cores_loc[loc_idx % len(cores_loc)]
                left = 0.15 + loc_idx * col_w2
                add_textbox(slide, f"📍 {local['nome']}", left, 0.75, col_w2-0.1, 0.45,
                            bold=True, size=12, bg=cor)

                sgs_aqui = [sg for sg in range(1, num_sg+1) if rodizio[sem_idx].get(sg) == loc_idx]
                texto = ""
                for sg in sgs_aqui:
                    texto += f"SG{sg}:\n"
                    for a in alunos_por_sg.get(sg, []):
                        texto += f"  {a['Nome Completo']}\n"
                add_textbox(slide, texto.strip(), left, 1.25, col_w2-0.1, 5.8,
                            size=8, bg=tuple(max(0,c-20) for c in cor), align=PP_ALIGN.LEFT)

        pptx_io = io.BytesIO()
        prs.save(pptx_io)
        pptx_io.seek(0)
        return pptx_io
    except ImportError:
        # Se python-pptx não estiver disponível, retorna bytes vazios
        return io.BytesIO(b"")
