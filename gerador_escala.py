"""
Motor de geração de escalas — backend do app Streamlit
Recebe a configuração do formulário e devolve Excel + CSV + auditoria
"""
from datetime import date, timedelta
from collections import defaultdict
import io, csv, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


def gerar_escala(config, alunos, subgrupos, locais, rodizio, regras, ch_limites):
    """
    Ponto de entrada principal.
    Retorna dict com: excel_bytes, csv_bytes, auditoria, metricas
    """
    # ── 1. Montar estrutura base ───────────────────────────────
    n_sem = config["n_semanas"]
    di_parts = config["data_inicio"].split("/")
    start_date = date(int(di_parts[2]), int(di_parts[1]), int(di_parts[0]))
    start_dates = {wk: start_date + timedelta(weeks=wk-1) for wk in range(1, n_sem+1)}

    # Montar lista de alunos com sg
    alunos_all = [(a["nome"], a["ra"], a["sg"]) for a in alunos]
    schedule = rodizio["schedule"]  # sg → {sem → abrev_local}

    # ── 2. Alocar turnos ──────────────────────────────────────
    na = defaultdict(set)  # nome → set de (ds, local_abrev, turno)

    # Função de carga horária por slot
    def h_slot(loc, t, diw):
        if loc not in regras:
            return 0
        r = regras[loc]
        if t == 'M':
            return r["manha"]["horas"]
        if t == 'F':
            return r["fds"]["h_manha"]
        if t == 'FT':
            return r["fds"]["h_tarde"]
        if t == 'R':
            if diw == 1:  # terça
                return r["tarde"]["terca_horas"]
            return r["tarde"].get("red_horas", r["tarde"]["horas"])
        if t == 'T':
            if diw == 1:
                return r["tarde"]["terca_horas"]
            return r["tarde"]["horas"]
        return 0

    def calc_h(nome, wk):
        ws_d = start_dates[wk]
        we_d = ws_d + timedelta(days=6)
        total = 0
        for ds, loc, t in na[nome]:
            d2 = date(int(ds[6:10]), int(ds[3:5]), int(ds[:2]))
            if ws_d <= d2 <= we_d:
                total += h_slot(loc, t, d2.weekday())
        return total

    def count_tarde(ds, loc):
        return sum(1 for nm, ra, sg in alunos_all
                   if (ds, loc, 'T') in na[nm] or (ds, loc, 'R') in na[nm])

    all_days = []
    d = start_date
    while d <= start_date + timedelta(weeks=n_sem) - timedelta(days=1):
        all_days.append(d)
        d += timedelta(days=1)

    # Etapa 2.1 — Manhãs
    for nome, ra, sg in alunos_all:
        for wk in range(1, n_sem+1):
            loc = schedule.get(sg, {}).get(wk)
            if not loc or loc not in regras: continue
            r = regras[loc]
            ws_d = start_dates[wk]
            for di in range(5):
                d2 = ws_d + timedelta(days=di)
                ds = d2.strftime('%d/%m/%Y')
                diw = d2.weekday()
                # Quinta: verificar regra
                if diw == 3 and r["manha"]["quinta"] == "Sem atividade":
                    continue
                na[nome].add((ds, loc, 'M'))

    # Etapa 2.2 — Tardes (rotativo por dia)
    # Agrupar alunos por par/bloco por semana
    pares_def = rodizio["pares_def"]
    for par_nome, par_sgs in pares_def.items():
        sem_locais = rodizio["pares"][par_nome]["semanas"]
        pool = [n for n, ra, sg in alunos_all if sg in par_sgs]
        n = len(pool)

        for sem_idx, local_nome in enumerate(sem_locais):
            wk = sem_idx + 1
            abrev = next((l["abrev"] for l in locais if l["nome"] == local_nome), None)
            if not abrev or abrev not in regras: continue
            r = regras[abrev]
            if not r["tarde"]["tem"]: continue

            ws_d = start_dates[wk]
            slots = r["tarde"]["slots"]
            off = sem_idx  # offset rotativo

            for di in range(5):
                d2 = ws_d + timedelta(days=di)
                ds = d2.strftime('%d/%m/%Y')
                diw = d2.weekday()

                # Quinta sem tarde
                if diw == 3:
                    if r["tarde"]["quinta"] == "Sem tarde (ENAMED/reunião)":
                        continue

                # Calcular quantos slots neste dia
                n_slots = slots
                if diw == 1:  # terça
                    if r["tarde"]["terca"] == "Sem tarde":
                        continue
                    if r["tarde"]["terca_todos"]:
                        # Todos fazem tarde na terça
                        for nm in pool:
                            na[nm].add((ds, abrev, 'R'))
                        continue

                # Distribuir slots com offset
                for slot in range(n_slots):
                    idx = (off + slot) % n
                    nm = pool[idx]
                    # Último slot = reduzido (R), outros = T
                    t_tipo = 'R' if (slot == n_slots - 1 and r["tarde"]["reduzido"]) else 'T'
                    if diw == 1:
                        t_tipo = 'R'
                    na[nm].add((ds, abrev, t_tipo))

    # Etapa 2.3 — FDS
    fds_days = [d for d in all_days if d.weekday() >= 5]
    for local in locais:
        abrev = local["abrev"]
        if abrev not in regras: continue
        r = regras[abrev]
        if not r["fds"]["tem"]: continue

        fds_r = r["fds"]
        n_m = fds_r["n_manha"]
        n_t = fds_r["n_tarde"]

        # Determinar pool elegível
        if fds_r["quem"] == "Alunos deste local":
            pool_origem = abrev
        else:
            pool_origem = fds_r.get("local_origem", abrev)

        # FDS manhã: n_m alunos por manhã, distribuídos rotatoriamente
        contagem_fds = defaultdict(int)
        for fds_d in fds_days:
            ds = fds_d.strftime('%d/%m/%Y')
            wk = next((w for w in range(1, n_sem+1)
                       if start_dates[w] <= fds_d <= start_dates[w]+timedelta(days=6)), None)
            if not wk: continue
            # Pool: alunos que estão no local origem nessa semana
            pool = [nm for nm, ra, sg in alunos_all
                    if schedule.get(sg, {}).get(wk) == pool_origem]
            if not pool: continue
            max_por = fds_r["max_por_aluno"]
            candidatos = sorted(pool, key=lambda nm: (contagem_fds[nm], pool.index(nm)))
            for i in range(min(n_m, len(pool))):
                escolhido = candidatos[i]
                if max_por == 0 or contagem_fds[escolhido] < max_por:
                    na[escolhido].add((ds, abrev, 'F'))
                    contagem_fds[escolhido] += 1

    # Etapa 2.4 — Ajuste de CH
    lim_abs_default = 43
    for nome, ra, sg in alunos_all:
        for wk in range(1, n_sem+1):
            loc = schedule.get(sg, {}).get(wk)
            if not loc or loc not in ch_limites: continue
            lim = ch_limites[loc]["com_fds"]
            lim_abs = ch_limites[loc]["absoluto"]
            ws_d = start_dates[wk]

            # Reduzir se passou do limite
            while calc_h(nome, wk) > lim_abs:
                # Remover tarde de menor impacto
                ws_d = start_dates[wk]
                we_d = ws_d + timedelta(days=6)
                tardes = [(ds, t) for ds, l, t in na[nome]
                          if l == loc and t in ['T', 'R']
                          and ws_d <= date(int(ds[6:10]), int(ds[3:5]), int(ds[:2])) <= we_d]
                if not tardes: break
                # Remover última tarde adicionada
                ds_rem, t_rem = tardes[-1]
                na[nome].discard((ds_rem, loc, t_rem))

    # ── 3. Gerar saídas ───────────────────────────────────────
    # CSV normalizado
    csv_rows = []
    for nome, ra, sg in alunos_all:
        for ds, loc, turno in sorted(na[nome]):
            d2 = date(int(ds[6:10]), int(ds[3:5]), int(ds[:2]))
            diw = d2.weekday()
            dia_sem = ['Seg','Ter','Qua','Qui','Sex','Sáb','Dom'][diw]
            wk = next((w for w in range(1, n_sem+1)
                       if start_dates[w] <= d2 <= start_dates[w]+timedelta(days=6)), 0)
            h = h_slot(loc, turno, diw)
            r = regras.get(loc, {})
            horario_map = {
                'M': f"{r.get('manha',{}).get('inicio','07:00')}-{r.get('manha',{}).get('fim','13:00')}",
                'T': f"{r.get('tarde',{}).get('inicio','12:00')}-{r.get('tarde',{}).get('fim','18:00')}",
                'R': f"{r.get('tarde',{}).get('red_inicio','12:00')}-{r.get('tarde',{}).get('red_fim','16:00')}",
                'F': "FDS manhã",
                'FT': "FDS tarde",
            }
            csv_rows.append({
                "disciplina": config["especialidade"],
                "ano": config["ano"],
                "grupo": config["grupo"],
                "turma": config["turma"],
                "sg": sg,
                "nome": nome,
                "ra": ra,
                "semana": wk,
                "data": ds,
                "dia_semana": dia_sem,
                "local": next((l["nome"] for l in locais if l["abrev"]==loc), loc),
                "local_abrev": loc,
                "turno": turno,
                "horario": horario_map.get(turno, ""),
                "horas": h,
                "fds": "S" if turno in ['F','FT'] else "N",
            })

    csv_buf = io.StringIO()
    if csv_rows:
        writer = csv.DictWriter(csv_buf, fieldnames=csv_rows[0].keys())
        writer.writeheader()
        writer.writerows(csv_rows)
    csv_bytes = csv_buf.getvalue().encode("utf-8")

    # ── Auditoria ──────────────────────────────────────────────
    erros_audit = []
    for nome, ra, sg in alunos_all:
        for wk in range(1, n_sem+1):
            loc = schedule.get(sg, {}).get(wk)
            if not loc or loc not in ch_limites: continue
            h = calc_h(nome, wk)
            lim_abs = ch_limites[loc]["absoluto"]
            if h > lim_abs:
                erros_audit.append(f"CH: {nome} S{wk}: {h}h > {lim_abs}h")

    audit_lines = [
        f"AUDITORIA — {config['especialidade']} {config['grupo']}/{config['turma']}",
        "=" * 60,
        f"Total turnos: {len(csv_rows)}",
        f"Erros encontrados: {len(erros_audit)}",
        "",
    ] + (erros_audit if erros_audit else ["✓ Zero erros"])
    auditoria = "\n".join(audit_lines).encode("utf-8")

    # ── Excel simples ──────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws_resumo = wb.active
    ws_resumo.title = "Resumo de Horas"

    # Cabeçalho
    headers = ["SG", "Nome", "RA"] + [f"S{w}" for w in range(1, n_sem+1)] + ["TOTAL"]
    for c, h in enumerate(headers, 1):
        cell = ws_resumo.cell(1, c, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color="1F4E79")
        cell.alignment = Alignment(horizontal="center")

    for row_idx, (nome, ra, sg) in enumerate(alunos_all, 2):
        ws_resumo.cell(row_idx, 1, f"SG{sg}")
        ws_resumo.cell(row_idx, 2, nome)
        ws_resumo.cell(row_idx, 3, ra)
        total = 0
        for wk in range(1, n_sem+1):
            h = calc_h(nome, wk)
            ws_resumo.cell(row_idx, 3+wk, f"{h}h")
            total += h
        ws_resumo.cell(row_idx, 3+n_sem+1, f"{total}h")

    excel_buf = io.BytesIO()
    wb.save(excel_buf)
    excel_bytes = excel_buf.getvalue()

    # ── Métricas ───────────────────────────────────────────────
    hs = []
    for nome, ra, sg in alunos_all:
        for wk in range(1, n_sem+1):
            loc = schedule.get(sg, {}).get(wk)
            if loc and loc in regras:
                hs.append(calc_h(nome, wk))

    metricas = {
        "total_turnos": len(csv_rows),
        "ch_media": f"{statistics.mean(hs):.0f}h" if hs else "-",
        "ch_min": f"{min(hs)}h" if hs else "-",
        "ch_max": f"{max(hs)}h" if hs else "-",
        "erros_auditoria": len(erros_audit),
    }

    return {
        "excel_bytes": excel_bytes,
        "csv_bytes": csv_bytes,
        "auditoria": auditoria,
        "metricas": metricas,
    }
