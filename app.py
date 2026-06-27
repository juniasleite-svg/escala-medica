import streamlit as st
import pandas as pd
import json
import datetime
import requests
import csv
import io

st.set_page_config(page_title="Gerador de Escalas Médicas", page_icon="🏥", layout="wide")

# ── API ──────────────────────────────────────────────────────────────────────
def _claude_http(api_key, mensagens, system_prompt="", max_tokens=8000):
    """Chamada crua à API — thread-safe (NÃO usa st.*). Retorna texto ou None."""
    if not api_key:
        return None
    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": api_key, "anthropic-version": "2023-06-01", "content-type": "application/json"},
            json={"model": "claude-sonnet-4-6", "max_tokens": max_tokens,
                  "system": system_prompt, "messages": mensagens},
            timeout=180
        )
        result = resp.json()
    except Exception:
        return None
    if "content" not in result:
        return None
    return result["content"][0]["text"]

def chamar_claude(mensagens, system_prompt="", max_tokens=8000):
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        st.error("Chave de API não configurada! Vá em Settings → Secrets.")
        return None
    texto = _claude_http(api_key, mensagens, system_prompt, max_tokens)
    if texto is None:
        st.error("Erro API: a IA não respondeu (timeout ou resposta inválida).")
    return texto

# ── JSON Parser robusto ───────────────────────────────────────────────────────
def extrair_json(texto):
    import re
    if not texto: return None

    # 1. ```json ... ```
    m = re.search(r'```json\s*([\s\S]*?)\s*```', texto)
    if m:
        try: return json.loads(m.group(1))
        except: pass

    # 2. ``` ... ```
    m = re.search(r'```\s*([\s\S]*?)\s*```', texto)
    if m:
        try: return json.loads(m.group(1))
        except: pass

    # 3. Remove comentários // e tenta
    linhas = [l for l in texto.split('\n') if not l.strip().startswith('//') and not l.strip().startswith('#')]
    txt = '\n'.join(linhas)

    # 4. Tenta parsear direto
    inicio = txt.find('{')
    if inicio >= 0:
        try: return json.loads(txt[inicio:])
        except: pass

        # 5. Encontra maior bloco JSON válido
        depth = 0
        for i, ch in enumerate(txt[inicio:]):
            if ch == '{': depth += 1
            elif ch == '}':
                depth -= 1
                if depth == 0:
                    try: return json.loads(txt[inicio:inicio+i+1])
                    except: pass

        # 6. JSON truncado — tenta completar
        parcial = txt[inicio:].strip()
        if parcial:
            # Conta chaves/colchetes abertos e tenta fechar
            try:
                stack = []
                in_str = False
                escape = False
                for ch in parcial:
                    if escape: escape = False; continue
                    if ch == '\\' and in_str: escape = True; continue
                    if ch == '"' and not escape: in_str = not in_str; continue
                    if not in_str:
                        if ch in '{[': stack.append(ch)
                        elif ch == '}' and stack and stack[-1] == '{': stack.pop()
                        elif ch == ']' and stack and stack[-1] == '[': stack.pop()

                # Fecha o que ficou aberto
                fechamento = ''
                for s in reversed(stack):
                    fechamento += '}' if s == '{' else ']'

                # Remove trailing comma/incomplete entry antes de fechar
                parcial_limpo = re.sub(r',\s*$', '', parcial.rstrip())
                # Remove entrada incompleta (linha que não fechou)
                linhas_json = parcial_limpo.split('\n')
                while linhas_json:
                    tentativa = '\n'.join(linhas_json) + fechamento
                    try:
                        resultado = json.loads(tentativa)
                        return resultado
                    except:
                        linhas_json.pop()  # Remove última linha e tenta novamente
            except: pass

    # 7. Entre primeiro { e último }
    try:
        s = texto.index('{'); e = texto.rindex('}') + 1
        return json.loads(texto[s:e])
    except: pass

    return None

# ── System prompt para geração ───────────────────────────────────────────────
SYSTEM_GERAR = """Você é um especialista em escalas de internato médico.
Gere a escala seguindo EXATAMENTE as regras do briefing.
Responda APENAS com JSON válido, sem texto antes ou depois, sem comentários //.

REGRAS FUNDAMENTAIS:
1. EXCLUSIVIDADE POR BLOCO: Cada bloco de rodízio pode ter vários serviços (ex: Bloco BP = Enf + PA Mandic). Um aluno SÓ pode estar em UM serviço do bloco por vez — nunca em dois serviços do mesmo bloco simultaneamente no mesmo dia/turno.
2. RODÍZIO INTERNO: Se o bloco tem 2 serviços e 4 SGs, 2 SGs vão ao Serviço 1 e 2 SGs vão ao Serviço 2. Eles NÃO se misturam no mesmo turno.
3. LIMITE DE CH SEMANAL — REGRA MAIS IMPORTANTE:
   - O limite é de 40h por semana por aluno (absoluto máximo: 43h).
   - Conte as horas de TODOS os turnos do dia: manhã (6h) + tarde (6h) + cinderela (4h) = 16h em UM dia.
   - Se um local tem manhã + tarde + cinderela 5 dias por semana = 80h → PROIBIDO.
   - Cinderela e turnos extras devem ser distribuídos para NÃO ultrapassar 40h/sem.
   - No resumo_horas, CALCULE as horas reais antes de preencher: some todos os turnos do SG na semana.
4. BLOQUEIOS: Aplique todos os bloqueios de manhã, tarde e FDS conforme configurado.

Estrutura obrigatória:
{
  "confirmacao": "resumo em 2 frases de como entendeu o rodízio e a distribuição",
  "calendario_rodizio": [
    {"semana": 1, "periodo": "06/07-12/07", "alocacao": {"SG1": "Bloco1-Srv1", "SG2": "Bloco1-Srv2", "SG3": "Bloco2"}}
  ],
  "resumo_horas": [
    {"sg": 1, "nome": "Nome Completo", "ra": "", "total_horas": 40, "semanas": [40,40,40,40,40,40,40,40], "plantoes_enf_fds": 1, "plantoes_ps_fds": 4, "cinderelas": 0}
  ],
  "auditoria": {"ch_ok": true, "cobertura_ok": true, "erros": [], "avisos": [], "aprovado": true}
}

NÃO inclua escala_detalhada nesta resposta — ela será gerada separadamente."""

SYSTEM_DETALHE = """Você é um especialista em escalas de internato médico.
Com base no briefing e no calendário de rodízio fornecido, gere APENAS a escala_detalhada.
Responda APENAS com JSON válido, sem texto antes ou depois.

═══════════ A REGRA QUE MAIS SE ERRA: CARGA HORÁRIA POR ALUNO ═══════════
O limite (veja o briefing, em geral 40h) é por SEMANA e por ALUNO INDIVIDUAL — não por serviço.
Um turno de manhã (~6h) + tarde (~6h) no MESMO dia = ~12h. Fazer isso 5 dias = 60h → PROIBIDO.
➡️ NUNCA coloque o mesmo aluno em manhã + tarde (+ cinderela) todos os dias da semana.

✅ COMO ACERTAR — REVEZAMENTO DE TURNOS ENTRE OS ALUNOS:
Os alunos de um mesmo serviço se DIVIDEM e se REVEZAM nos turnos, para que CADA UM fique ≤ limite.
Exemplo (4 alunos A,B,C,D num serviço com manhã e tarde):
   Seg: Manhã=[A,B]  Tarde=[C,D]
   Ter: Manhã=[C,D]  Tarde=[A,B]
   Qua: Manhã=[A,B]  Tarde=[C,D]
   Qui: Manhã=[C,D]  Tarde=[A,B]
   Sex: Manhã=[A,B]  Tarde=[C,D]
Assim cada aluno faz ~metade dos turnos e fica dentro do limite. NÃO repita o mesmo aluno em todos os turnos.

═══════════ QUANTIDADE DE ALUNOS POR TURNO (mín/máx) ═══════════
Cada serviço tem mín e máx de alunos por turno/dia (no briefing). Coloque a quantidade DENTRO dessa faixa
em CADA turno de CADA dia. Se a Enfermaria pede 2 por manhã, coloque exatamente 2 (não 4).

═══════════ BLOCO COM VÁRIOS SERVIÇOS (rodízio interno) ═══════════
Quando um bloco tem 2+ serviços (ex: Bloco Pediátrico = Enfermaria + PA Mandic):
- DIVIDA os alunos do bloco ENTRE os serviços ao mesmo tempo (ex: parte na Enf, parte no PA).
- NENHUM serviço do bloco pode ficar VAZIO numa semana em que o bloco está ativo.
- Ao longo das semanas, TROQUE: quem estava na Enf vai pro PA e vice-versa.
- Um aluno nunca aparece em 2 serviços/locais no mesmo dia/turno.

═══════════ COBERTURA E BLOQUEIOS ═══════════
1. Para CADA serviço ativo, gere entrada para CADA turno em CADA dia útil (Seg–Sex), exceto bloqueios.
2. Só omita um turno se houver bloqueio explícito ("Sem tarde na quinta" → só quinta tarde).
   "Terça tarde 12-16h" → gere a tarde de terça reduzida, NÃO omita.

FORMATO: 1 entrada por (semana + data + local + turno), listando os alunos daquele turno.
   - Datas: DD/MM ; horas: duração do turno (manhã 6h, tarde 6h, cinderela 4h).

⚠️ ANTES DE RESPONDER: para CADA aluno, some as horas da semana. Se algum passar do limite, redistribua
os turnos entre os colegas do subgrupo até todos ficarem ≤ limite.

Formato:
{
  "escala_detalhada": [
    {"semana": 1, "data": "06/07", "dia": "Seg", "local": "Enfermaria", "turno": "Manhã", "horario": "07-13h", "horas": 6, "sg": 1, "alunos": ["Nome1","Nome2"]},
    {"semana": 1, "data": "06/07", "dia": "Seg", "local": "Enfermaria", "turno": "Tarde", "horario": "13-19h", "horas": 6, "sg": 1, "alunos": ["Nome3","Nome4"]}
  ]
}"""

SYSTEM_CORRIGIR = """Você é um especialista em escalas de internato médico.
Recebeu as entradas de UMA SEMANA de uma escala que VIOLA o limite de carga horária semanal
e/ou tem aluno em dois locais ao mesmo tempo. Reescreva as entradas dessa semana corrigindo
TODAS as violações, SEM perder cobertura de nenhum local/turno.

COMO CORRIGIR (essencial):
- Um aluno NÃO pode fazer manhã + tarde + cinderela todos os dias — isso estoura a carga horária.
- DISTRIBUA (revezamento) os turnos ENTRE os alunos do mesmo subgrupo/serviço: parte do grupo
  cobre a manhã, outra parte cobre a tarde, alternando ao longo da semana, de forma que CADA
  aluno fique com no máximo {limite}h na semana.
- TODOS os locais/turnos que já apareciam devem continuar cobertos — não deixe nenhum vazio.
- Um aluno só pode estar em UM local por dia/turno.
- Não tire ninguém do seu bloco/serviço — apenas redistribua os turnos entre os alunos.
- Mantenha as mesmas datas, dias, locais, horários e o campo "horas" de cada turno.

Responda APENAS com JSON válido, sem texto antes ou depois:
{{"escala_detalhada": [ ...todas as entradas corrigidas SOMENTE desta semana... ]}}"""

# ── Validação determinística — o Python é o juiz (não a IA) ───────────────────
import re as _re_val
from collections import defaultdict as _dd

DIAS_UTEIS = ["Seg", "Ter", "Qua", "Qui", "Sex"]

def _dia_curto(d):
    d = str(d or "").strip()
    dl = d.lower()
    mapa = {"segunda": "Seg", "terça": "Ter", "terca": "Ter", "quarta": "Qua",
            "quinta": "Qui", "sexta": "Sex", "sábado": "Sab", "sabado": "Sab", "domingo": "Dom"}
    for k, v in mapa.items():
        if dl.startswith(k):
            return v
    if d[:3].capitalize() in ("Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"):
        return d[:3].capitalize()
    return d

def _turno_key(t):
    t = str(t or "").lower()
    if "manh" in t:
        return "manha"
    if "tard" in t:
        return "tarde"
    if "cind" in t or "cinder" in t:
        return "cind"
    if "fds" in t:
        return "fds"
    return t.strip()

def _horas_entrada(e):
    """Horas reais de um turno: usa o campo 'horas', senão deriva do horário, senão do turno."""
    try:
        h = float(e.get("horas"))
        if h > 0:
            return h
    except (TypeError, ValueError):
        pass
    nums = _re_val.findall(r"(\d{1,2})", str(e.get("horario", "")))
    if len(nums) >= 2:
        try:
            ini, fim = int(nums[0]), int(nums[1])
            if fim > ini:
                return float(fim - ini)
        except ValueError:
            pass
    return {"manha": 6.0, "tarde": 6.0, "cind": 4.0, "fds": 6.0}.get(_turno_key(e.get("turno")), 0.0)

def _alunos_entrada(e):
    a = e.get("alunos")
    if isinstance(a, list) and a:
        return [str(x).strip() for x in a if str(x).strip()]
    nome = e.get("nome")
    if isinstance(nome, str) and nome.strip():
        return [nome.strip()]
    return []

def _bloqueios_map(config):
    """Mapa de bloqueios -> {turno_key: {(local_lower, dia_curto)}} para não acusar buraco onde há bloqueio."""
    blk = _dd(set)
    for loc in config.get("locais", []):
        servs = [loc] + (loc.get("servicos_extras") or [])
        for s in servs:
            nomes = {str(s.get("nome", "")).strip().lower(),
                     str(s.get("abrev", "")).strip().lower(),
                     str(loc.get("nome_bloco", "")).strip().lower()}
            nomes = {n for n in nomes if n}
            for b in (s.get("bloqueios_manha") or []):
                if str(b.get("tipo", "")).lower().startswith("sem"):
                    for n in nomes:
                        blk["manha"].add((n, _dia_curto(b.get("dia"))))
            for b in (s.get("bloqueios_tarde") or []):
                if str(b.get("tipo", "")).lower().startswith("sem"):
                    for n in nomes:
                        blk["tarde"].add((n, _dia_curto(b.get("dia"))))
    return blk

def _norm(s):
    return str(s or "").strip().lower()

def _nome_bate(nomes_set, local_lower):
    """Casa um conjunto de nomes/abreviações de serviço com o nome do local da entrada."""
    if local_lower in nomes_set:
        return True
    return any(n and (n in local_lower or local_lower in n) for n in nomes_set)

def _servicos_config(config):
    """Lista de serviços a partir do config: nomes, bloco e (min,max) por turno."""
    servs = []
    for loc in config.get("locais", []):
        bloco = loc.get("nome_bloco") or loc.get("nome") or ""
        irmaos = [loc] + (loc.get("servicos_extras") or [])
        for s in irmaos:
            nomes = {_norm(s.get("nome")), _norm(s.get("abrev"))} - {""}
            if not nomes:
                continue
            turnos = {}
            for tk, (mn, mx, ativo) in {
                "manha": (s.get("min_manha"), s.get("max_manha"), s.get("manha")),
                "tarde": (s.get("min_tarde"), s.get("max_tarde"), s.get("tarde")),
                "cind":  (s.get("min_cind"),  s.get("max_cind"),  s.get("cinderela")),
            }.items():
                if not ativo:
                    continue
                try: mn = int(mn or 0)
                except (TypeError, ValueError): mn = 0
                try: mx = int(mx) if mx not in (None, "", 0) else None
                except (TypeError, ValueError): mx = None
                turnos[tk] = (mn, mx)
            servs.append({"nomes": nomes, "bloco": bloco,
                          "label": s.get("nome") or s.get("abrev") or "", "turnos": turnos})
    return servs

def validar_escala(dados, config):
    """Confere a escala de verdade: horas/semana, conflitos, cobertura, mín/máx por serviço e rodízio dentro do bloco."""
    det = dados.get("escala_detalhada") or []
    reg = config.get("regras_especiais", {})
    limite = int(reg.get("limite_ch", 40))
    limite_abs = int(reg.get("limite_abs", 43))
    regra_quinta = str(reg.get("quinta", "")).lower()

    horas = _dd(float)   # (aluno, semana) -> horas
    ocup = _dd(set)      # (aluno, semana, dia, turno) -> {locais}
    cob = _dd(set)       # (local, turno, semana) -> {dias}
    qtd = _dd(int)       # (local, turno, semana, dia) -> nº alunos

    for e in det:
        sem = e.get("semana", "?")
        dia = _dia_curto(e.get("dia") or e.get("data"))
        turno = _turno_key(e.get("turno"))
        local = str(e.get("local", "?"))
        h = _horas_entrada(e)
        als = _alunos_entrada(e)
        for al in als:
            horas[(al, sem)] += h
            ocup[(al, sem, dia, turno)].add(local)
        cob[(local, turno, sem)].add(dia)
        qtd[(local, turno, sem, dia)] += len(als)

    estouros = []
    for (al, sem), h in horas.items():
        if h > limite_abs:
            estouros.append({"aluno": al, "semana": sem, "horas": round(h, 1), "limite": limite_abs, "nivel": "absoluto"})
        elif h > limite:
            estouros.append({"aluno": al, "semana": sem, "horas": round(h, 1), "limite": limite, "nivel": "padrao"})
    estouros.sort(key=lambda x: -x["horas"])

    conflitos = []
    for (al, sem, dia, turno), locs in ocup.items():
        if len(locs) > 1:
            conflitos.append({"aluno": al, "semana": sem, "dia": dia, "turno": turno, "locais": sorted(locs)})

    blk = _bloqueios_map(config)
    buracos = []
    for (local, turno, sem), dias in cob.items():
        if turno in ("cind", "fds"):
            continue
        ll = local.lower()
        falt = []
        for d in DIAS_UTEIS:
            if d in dias:
                continue
            if (ll, d) in blk.get(turno, set()):
                continue
            if d == "Qui" and turno == "tarde" and "sem tarde" in regra_quinta:
                continue
            falt.append(d)
        if falt:
            buracos.append({"local": local, "turno": turno, "semana": sem, "dias": falt})

    # ── Mín/máx de alunos por serviço/turno/dia ──────────────────────────────
    servs_cfg = _servicos_config(config)
    desvios = []
    for (local, turno, sem, dia), n in qtd.items():
        srv = next((s for s in servs_cfg if turno in s["turnos"] and _nome_bate(s["nomes"], _norm(local))), None)
        if not srv:
            continue
        mn, mx = srv["turnos"][turno]
        if n < mn:
            desvios.append({"local": local, "turno": turno, "semana": sem, "dia": dia, "qtd": n, "min": mn, "max": mx, "tipo": "abaixo"})
        elif mx is not None and n > mx:
            desvios.append({"local": local, "turno": turno, "semana": sem, "dia": dia, "qtd": n, "min": mn, "max": mx, "tipo": "acima"})

    # ── Rodízio dentro do bloco: serviço do bloco sem ninguém numa semana ativa ─
    presentes_sem = _dd(set)  # semana -> {local_lower presentes}
    for (local, turno, sem) in cob.keys():
        presentes_sem[sem].add(_norm(local))
    ausentes = []
    for loc in config.get("locais", []):
        irmaos = [loc] + (loc.get("servicos_extras") or [])
        if len(irmaos) < 2:
            continue
        bloco = loc.get("nome_bloco") or loc.get("nome") or ""
        nomes_irmaos = [({_norm(s.get("nome")), _norm(s.get("abrev"))} - {""}) for s in irmaos]
        for sem, presentes in presentes_sem.items():
            ativos = [i for i, nm in enumerate(nomes_irmaos) if nm and any(_nome_bate(nm, p) for p in presentes)]
            if not ativos:
                continue  # bloco não está ativo nesta semana
            for i, s in enumerate(irmaos):
                nm = nomes_irmaos[i]
                if nm and not any(_nome_bate(nm, p) for p in presentes):
                    ausentes.append({"servico": s.get("nome") or s.get("abrev") or f"serviço {i+1}",
                                     "bloco": bloco, "semana": sem})

    semanas_ruins = sorted(
        {e["semana"] for e in estouros} | {c["semana"] for c in conflitos}
        | {d["semana"] for d in desvios} | {a["semana"] for a in ausentes},
        key=lambda x: int(x) if str(x).isdigit() else 999)
    return {
        "estouros": estouros, "conflitos": conflitos, "buracos": buracos,
        "desvios": desvios, "ausentes": ausentes,
        "limite": limite, "limite_abs": limite_abs,
        "ok": (not estouros and not conflitos and not desvios and not ausentes),
        "semanas_ruins": semanas_ruins,
    }

def recalcular_resumo_horas(dados, config):
    """Recalcula o resumo de horas a partir da escala_detalhada real (não confia no que a IA declarou)."""
    det = dados.get("escala_detalhada") or []
    alunos_sg = config.get("alunos_por_sg", {})
    aluno_sg = {}
    for sg, nomes in alunos_sg.items():
        for n in nomes:
            aluno_sg[str(n).strip()] = sg
    num_sem = int(config.get("num_semanas", 8))
    hs = _dd(lambda: _dd(float))
    cind = _dd(int)
    for e in det:
        h = _horas_entrada(e)
        tk = _turno_key(e.get("turno"))
        try:
            sem = int(e.get("semana"))
        except (TypeError, ValueError):
            continue
        for al in _alunos_entrada(e):
            hs[al][sem] += h
            if tk == "cind":
                cind[al] += 1
    rh_old = {r.get("nome", ""): r for r in (dados.get("resumo_horas") or [])}
    nomes_todos = list(aluno_sg.keys()) or list(hs.keys())
    linhas = []
    for al in nomes_todos:
        semanas = [round(hs[al].get(i + 1, 0.0), 1) for i in range(num_sem)]
        base = dict(rh_old.get(al, {}))
        base.update({
            "sg": aluno_sg.get(al, "") or base.get("sg", ""),
            "nome": al, "ra": base.get("ra", ""),
            "total_horas": round(sum(semanas), 1),
            "semanas": semanas,
            "cinderelas": cind.get(al, base.get("cinderelas", 0)),
        })
        linhas.append(base)

    def _sg_ord(r):
        s = str(r.get("sg", "")).strip()
        return (int(s) if s.isdigit() else 99, r.get("nome", ""))
    linhas.sort(key=_sg_ord)
    return linhas

def corrigir_escala_loop(dados, config, briefing="", max_rodadas=2):
    """Manda os erros concretos de volta pra IA, semana a semana (TODAS em paralelo), até ficar válida."""
    import concurrent.futures
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    reg = config.get("regras_especiais", {})
    ctx = (
        f"Subgrupos e alunos:\n{json.dumps(config.get('alunos_por_sg', {}), ensure_ascii=False)}\n\n"
        f"Regras: Limite {reg.get('limite_ch', 40)}h/semana por aluno (absoluto {reg.get('limite_abs', 43)}h). "
        f"Quinta: {reg.get('quinta', '')}. Terça: {reg.get('terca', '')}. FDS: {reg.get('fds', '')}."
    )
    for _ in range(max_rodadas):
        val = validar_escala(dados, config)
        if val["ok"]:
            break
        det = dados.get("escala_detalhada") or []
        tarefas = {}  # semana -> mensagem de correção
        for sem in val["semanas_ruins"]:
            entradas_sem = [e for e in det if e.get("semana") == sem]
            est = [e for e in val["estouros"] if e["semana"] == sem]
            con = [c for c in val["conflitos"] if c["semana"] == sem]
            des = [d for d in val.get("desvios", []) if d["semana"] == sem]
            aus = [a for a in val.get("ausentes", []) if a["semana"] == sem]
            if not entradas_sem or (not est and not con and not des and not aus):
                continue
            problemas = []
            for e in est:
                problemas.append(f"- {e['aluno']}: {e['horas']}h nesta semana (limite {val['limite']}h) — reduza para no máximo {val['limite']}h")
            for c in con:
                problemas.append(f"- {c['aluno']}: em 2 locais no mesmo {c['dia']}/{c['turno']} ({', '.join(c['locais'])}) — deixe em apenas um")
            for d in des:
                faixa = f"entre {d['min']} e {d['max']}" if d['max'] is not None else f"no mínimo {d['min']}"
                problemas.append(f"- {d['local']}/{d['turno']} em {d['dia']}: tem {d['qtd']} alunos — ajuste para {faixa}")
            aus_vistos = set()
            for a in aus:
                if a["servico"] in aus_vistos:
                    continue
                aus_vistos.add(a["servico"])
                problemas.append(f"- O serviço '{a['servico']}' (bloco {a['bloco']}) ficou SEM ninguém — divida os alunos do bloco entre os serviços (revezamento), parte aqui e parte nos outros")
            tarefas[sem] = (
                f"{ctx}\n\nSEMANA {sem} — VIOLAÇÕES A CORRIGIR:\n" + "\n".join(problemas) +
                f"\n\nENTRADAS ATUAIS DESTA SEMANA (corrija e devolva TODAS):\n" +
                json.dumps(entradas_sem, ensure_ascii=False)
            )
        if not tarefas:
            break
        sysp = SYSTEM_CORRIGIR.format(limite=val["limite"])
        novas_por_sem = {}
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(8, len(tarefas))) as ex:
            futuros = {ex.submit(_claude_http, api_key, [{"role": "user", "content": tarefas[sem]}], sysp, 8000): sem
                       for sem in tarefas}
            for fut in concurrent.futures.as_completed(futuros):
                sem = futuros[fut]
                novas = (extrair_json(fut.result()) or {}).get("escala_detalhada")
                if novas:
                    for e in novas:
                        e["semana"] = sem
                    novas_por_sem[sem] = novas
        if not novas_por_sem:
            break
        for sem, novas in novas_por_sem.items():
            det = [e for e in det if e.get("semana") != sem] + novas
        dados["escala_detalhada"] = det
        dados["resumo_horas"] = recalcular_resumo_horas(dados, config)

    try:
        dados["escala_detalhada"].sort(
            key=lambda e: (int(e.get("semana", 0)) if str(e.get("semana", "")).isdigit() else 0, str(e.get("data", "")))
        )
    except Exception:
        pass
    dados["resumo_horas"] = recalcular_resumo_horas(dados, config)
    return dados, validar_escala(dados, config)

def gerar_detalhada_por_semana(briefing, calendario_json, alunos_por_sg, config, num_semanas, ui=True):
    """Gera a escala_detalhada por semana, TODAS em paralelo (rápido) e sem truncar a resposta."""
    import concurrent.futures
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    limite = config.get("regras_especiais", {}).get("limite_ch", 40)
    n = max(int(num_semanas), 1)

    def _msg(sem):
        return (
            f"Briefing:\n{briefing}\n\n"
            f"Calendário de rodízio (todas as semanas):\n{calendario_json}\n\n"
            f"Alunos por subgrupo:\n{json.dumps(alunos_por_sg, ensure_ascii=False)}\n\n"
            f"Gere a escala_detalhada APENAS da SEMANA {sem}: todos os locais ativos nessa semana, "
            f"todos os turnos em todos os dias úteis (e FDS se houver), com os alunos REVEZADOS para "
            f"ninguém passar de {limite}h. Respeite o mín/máx por turno e divida os alunos entre os "
            f"serviços de cada bloco. NÃO gere outras semanas."
        )

    resultados, feitos = {}, 0
    barra = st.progress(0.0, text="Gerando escala detalhada (semanas em paralelo)...") if ui else None
    with concurrent.futures.ThreadPoolExecutor(max_workers=min(8, n)) as ex:
        futuros = {ex.submit(_claude_http, api_key, [{"role": "user", "content": _msg(sem)}],
                             SYSTEM_DETALHE, 8000): sem for sem in range(1, n + 1)}
        for fut in concurrent.futures.as_completed(futuros):
            sem = futuros[fut]
            novas = (extrair_json(fut.result()) or {}).get("escala_detalhada")
            if novas:
                for e in novas:
                    e["semana"] = sem
                resultados[sem] = novas
            feitos += 1
            if barra:
                barra.progress(feitos / n, text=f"Escala detalhada: {feitos}/{n} semanas prontas")
    if barra:
        barra.empty()
    todas = []
    for sem in range(1, n + 1):
        todas.extend(resultados.get(sem, []))
    return todas

def mostrar_validacao(val):
    """Mostra o resultado da validação real na tela."""
    if val["ok"] and not val["buracos"]:
        st.success(f"✅ Validação do sistema: todos os alunos ≤ {val['limite']}h/semana, ninguém em 2 lugares ao mesmo tempo e cobertura completa.")
        return
    if val["ok"]:
        st.success(f"✅ Carga horária OK (todos ≤ {val['limite']}h/sem) e sem alunos em 2 lugares ao mesmo tempo.")
    if val["estouros"]:
        st.error(f"❌ **{len(val['estouros'])} aluno(s)/semana acima do limite de carga horária:**")
        for e in val["estouros"][:30]:
            st.markdown(f"- **{e['aluno']}** — semana {e['semana']}: **{e['horas']}h** (limite {e['limite']}h)")
        if len(val["estouros"]) > 30:
            st.caption(f"...e mais {len(val['estouros']) - 30}")
    if val["conflitos"]:
        st.error(f"❌ **{len(val['conflitos'])} caso(s) de aluno em 2 locais ao mesmo tempo:**")
        for c in val["conflitos"][:30]:
            st.markdown(f"- **{c['aluno']}** — sem {c['semana']}, {c['dia']} {c['turno']}: {', '.join(c['locais'])}")
    if val.get("desvios"):
        acima = [d for d in val["desvios"] if d["tipo"] == "acima"]
        abaixo = [d for d in val["desvios"] if d["tipo"] == "abaixo"]
        st.error(f"❌ **{len(val['desvios'])} caso(s) com número de alunos fora do configurado por serviço/turno:**")
        for d in (acima + abaixo)[:30]:
            faixa = f"{d['min']}–{d['max']}" if d['max'] is not None else f"mín {d['min']}"
            st.markdown(f"- {d['local']} / {d['turno']} — sem {d['semana']}, {d['dia']}: **{d['qtd']} alunos** (esperado {faixa})")
    if val.get("ausentes"):
        st.error("❌ **Serviços de um bloco que ficaram sem ninguém (falta rodízio interno):**")
        vistos = set()
        for a in val["ausentes"]:
            chave = (a["servico"], a["semana"])
            if chave in vistos:
                continue
            vistos.add(chave)
            st.markdown(f"- **{a['servico']}** (bloco {a['bloco']}) — sem {a['semana']}: ninguém alocado; distribua parte dos alunos do bloco para cá")
    if val["buracos"]:
        st.warning("⚠️ **Possíveis dias sem cobertura** (confira se não é bloqueio legítimo):")
        for b in val["buracos"][:30]:
            st.markdown(f"- {b['local']} / {b['turno']} — sem {b['semana']}: faltam {', '.join(b['dias'])}")

# ── Mostrar resultado ────────────────────────────────────────────────────────
def mostrar_resultado(resposta_raw, esp, grupo, turma):
    dados = extrair_json(resposta_raw)
    if dados is None:
        st.subheader("📄 Resposta da IA")
        st.text_area("Conteúdo gerado", resposta_raw, height=300)
        st.warning("A IA respondeu em formato inesperado. Verifique o conteúdo acima.")
        return

    # ── Validação REAL (o sistema confere de verdade, não a IA) ──────────────
    config_v = st.session_state.get("config_atual", {})
    if dados.get("escala_detalhada"):
        dados["resumo_horas"] = recalcular_resumo_horas(dados, config_v)
        val = validar_escala(dados, config_v)
        st.subheader("⚖️ Validação automática (conferida pelo sistema)")
        mostrar_validacao(val)
        if not val["ok"]:
            if st.button("🔧 Rebalancear automaticamente (respeitar 40h/sem)", type="primary", key="btn_rebal"):
                with st.spinner("Rebalanceando turnos entre os alunos dos subgrupos... ⏳"):
                    novo, _ = corrigir_escala_loop(dados, config_v, st.session_state.get("briefing_atual", ""))
                st.session_state.escala_gerada = json.dumps(novo, ensure_ascii=False)
                st.rerun()
        st.divider()

    # Confirmação
    if dados.get("confirmacao"):
        with st.expander("📋 O que a IA entendeu", expanded=False):
            st.write(dados["confirmacao"])

    # Observações auto-reportadas pela IA (informativo — a validação acima é a que vale)
    audit = dados.get("auditoria", {})
    if audit.get("erros") or audit.get("avisos"):
        with st.expander("🤖 Observações da IA (informativo)", expanded=False):
            for err in audit.get("erros", []):
                st.write(f"❌ {err}")
            for av in audit.get("avisos", []):
                st.write(f"⚠️ {av}")

    # Calendário
    if dados.get("calendario_rodizio"):
        st.subheader("📅 Calendário de Rodízio")
        cal = dados["calendario_rodizio"]
        rows = []
        for r in cal:
            row = {"Semana": f"Sem {r.get('semana','')} | {r.get('periodo','')}"}
            row.update(r.get("alocacao", {}))
            rows.append(row)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # Resumo de horas
    if dados.get("resumo_horas"):
        st.subheader("⏱️ Resumo de Horas")
        df_h = pd.DataFrame(dados["resumo_horas"])
        if "sg" in df_h.columns: df_h = df_h.sort_values(["sg","nome"])
        st.dataframe(df_h, use_container_width=True, hide_index=True)

    # Downloads
    st.subheader("📥 Exportar")
    col1, col2, col3 = st.columns(3)

    # Excel
    try:
        from gerador_excel_final import gerar_excel_completo
        cfg = st.session_state.get("config_atual", {})
        cfg.update({"especialidade": esp, "grupo": grupo, "turma": turma})
        xlsx = gerar_excel_completo(dados, cfg)
        with col1:
            st.download_button("📊 Excel (9 abas)", data=xlsx,
                file_name=f"Escala_{esp}_{grupo}_{turma}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
    except Exception as e:
        with col1:
            st.warning(f"Erro Excel: {e}")

    # CSV nominal
    if dados.get("escala_detalhada"):
        buf = io.StringIO()
        fields = ["semana","data","dia","local","turno","horario","horas","sg","nome","ra"]
        w = csv.DictWriter(buf, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for entry in dados["escala_detalhada"]:
            alunos = entry.get("alunos", [])
            nome = entry.get("nome","") if isinstance(entry.get("nome"), str) else ""
            if alunos and not nome:
                for a in alunos:
                    row = {k: entry.get(k,"") for k in fields}
                    row["nome"] = a
                    w.writerow(row)
            else:
                w.writerow({k: entry.get(k,"") for k in fields})
        with col2:
            st.download_button("📄 CSV — Escala detalhada",
                data=buf.getvalue().encode("utf-8-sig"),
                file_name=f"Escala_{esp}_{grupo}_{turma}_nominal.csv",
                mime="text/csv", use_container_width=True)

    # CSV horas
    if dados.get("resumo_horas"):
        buf2 = io.StringIO()
        rh = dados["resumo_horas"]
        if rh:
            f2 = list(rh[0].keys())
            w2 = csv.DictWriter(buf2, fieldnames=f2, extrasaction="ignore")
            w2.writeheader()
            for r in rh:
                row2 = {k: " | ".join(str(x) for x in v) if isinstance(v, list) else v for k,v in r.items()}
                w2.writerow(row2)
        with col3:
            st.download_button("📄 CSV — Resumo horas",
                data=buf2.getvalue().encode("utf-8-sig"),
                file_name=f"Escala_{esp}_{grupo}_{turma}_horas.csv",
                mime="text/csv", use_container_width=True)

    # Correção
    st.divider()
    with st.expander("🔁 Solicitar correção à IA"):
        correcao = st.text_area("Descreva o que precisa corrigir:", key="correcao_txt",
            placeholder="Ex: A tarde de quinta deve ser bloqueada em todos os locais\nAdicionar feriado em 09/07")
        if st.button("Enviar correção", key="btn_correcao"):
            with st.spinner("IA corrigindo..."):
                r2 = chamar_claude([
                    {"role": "user", "content": f"Briefing original:\n{st.session_state.get('briefing_atual','')}"},
                    {"role": "assistant", "content": resposta_raw},
                    {"role": "user", "content": f"Corrija:\n{correcao}\nRetorne JSON completo corrigido."}
                ], system_prompt=SYSTEM_GERAR)
                if r2:
                    st.session_state.escala_gerada = r2
                    st.rerun()

# ════════════════════════════════════════════════════════════════════════════
# TOPO
# ════════════════════════════════════════════════════════════════════════════
st.title("🏥 Gerador de Escalas Médicas")
st.caption("Powered by Claude AI — geração inteligente com todas as regras da sua especialidade")

modo = st.radio("Como deseja começar?",
    ["✨ Criar nova escala", "📂 Importar escala existente e modificar"],
    horizontal=True)
st.divider()

# ════════════════════════════════════════════════════════════════════════════
# MODO IMPORTAR — só faz o upload e análise, depois redireciona pro formulário
# ════════════════════════════════════════════════════════════════════════════
if modo == "📂 Importar escala existente e modificar":
    st.header("📂 Importar Escala Existente")
    st.caption("A IA lê seu Excel e preenche o formulário automaticamente — você edita o que quiser e gera.")

    col_up1, col_up2 = st.columns(2)
    with col_up1:
        arquivo_escala = st.file_uploader("Upload da escala existente (.xlsx)", type=["xlsx"], key="imp_escala")
    with col_up2:
        arquivo_alunos_imp = st.file_uploader("Base de alunos (.xlsx — opcional)", type=["xlsx"], key="imp_alunos_file")

    if arquivo_escala:
        if st.button("🤖 Analisar e preencher formulário com IA", type="primary", use_container_width=True):
            with st.spinner("Lendo e interpretando sua escala..."):
                _bytes_escala = arquivo_escala.read()
                xls = pd.ExcelFile(io.BytesIO(_bytes_escala), engine="openpyxl")
                conteudo = ""
                for sheet in xls.sheet_names:
                    df = pd.read_excel(io.BytesIO(_bytes_escala), engine="openpyxl", sheet_name=sheet, header=None)
                    conteudo += f"\n\n=== ABA: {sheet} ===\n{df.fillna('').to_string(index=False, header=False)}"
                    if len(conteudo) > 12000:
                        conteudo = conteudo[:12000] + "\n...[truncado]"
                        break

                prompt = f"""Analise esta escala médica e extraia o briefing completo em JSON.

CONTEÚDO:
{conteudo}

Retorne APENAS JSON válido (sem markdown, sem comentários) com esta estrutura:
{{
  "especialidade": "",
  "grupo": "",
  "turma": "",
  "ano_curso": "4º Ano",
  "data_inicio": "YYYY-MM-DD",
  "num_semanas": 8,
  "num_sg": 6,
  "alunos_por_sg": {{"1": ["Nome1"], "2": ["Nome2"]}},
  "num_locais": 3,
  "locais": [
    {{
      "nome": "", "abrev": "", "obs": "",
      "manha": "07-13h", "min_manha": 0, "max_manha": 6,
      "tarde": "12-18h", "min_tarde": 3, "max_tarde": 4,
      "bloqueios_tarde": [{{"dia": "Qui", "tipo": "Sem tarde", "horario": ""}}, {{"dia": "Ter", "tipo": "Horário reduzido", "horario": "12-16h"}}],
      "cinderela": "", "min_cind": 0, "max_cind": 0, "dias_cind": [],
      "fds": false, "fds_manha": "", "fds_tarde": "", "fds_cind": "",
      "fds_min_manha": 0, "fds_max_manha": 2, "fds_min_tarde": 0, "fds_max_tarde": 2,
      "fds_min_cind": 0, "fds_max_cind": 0,
      "fds_quem": "", "fds_comp": "", "servico2": "", "servico2_obs": "",
      "duracao_por_sg": {{}}, "cov_tarde": 3
    }}
  ],
  "rodizio_desc": "descrição do rodízio",
  "regra_quinta": "Sem tarde (ENAMED para todos)",
  "regra_terca": "Tarde encurtada 12-16h",
  "limite_ch": 40,
  "limite_abs": 43,
  "regra_fds": "",
  "regras_extras": "",
  "resumo": "descrição em português"
}}"""

                resposta = chamar_claude([{"role": "user", "content": prompt}], max_tokens=6000)
                if resposta:
                    dados = extrair_json(resposta)
                    if dados:
                        # Salvar no session_state para pré-preencher o formulário
                        st.session_state.prefill = dados
                        # Se tiver base de alunos, tentar carregar também
                        if arquivo_alunos_imp:
                            st.session_state.prefill["_arquivo_alunos"] = True
                        st.success("✅ Escala analisada! Role para baixo para ver o formulário pré-preenchido.")
                        st.rerun()
                    else:
                        st.error("Não consegui extrair os dados. Tente novamente.")

    if "prefill" not in st.session_state:
        st.info("👆 Faça o upload da sua escala e clique em analisar para preencher o formulário automaticamente.")
        st.stop()
    else:
        st.success("✅ Formulário preenchido! Revise e edite o que precisar abaixo.")
        st.divider()

# ════════════════════════════════════════════════════════════════════════════
# FORMULÁRIO ÚNICO (usado por ambos os modos)
# ════════════════════════════════════════════════════════════════════════════

# Pegar dados pré-preenchidos se existirem
pf = st.session_state.get("prefill", {})

if pf:
    st.info(f"💡 **Importado:** {pf.get('resumo','')}")

st.header("📋 Briefing da Escala")
st.caption("Preencha com atenção — quanto mais detalhado, mais precisa a escala gerada.")

# BLOCO 1
with st.expander("📌 Bloco 1 — Identificação", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        especialidade = st.text_input("Especialidade *",
            value=pf.get("especialidade",""),
            placeholder="ex: Clínica Médica")
        anos = ["3º Ano","4º Ano","5º Ano","6º Ano"]
        ano_idx = anos.index(pf.get("ano_curso","4º Ano")) if pf.get("ano_curso") in anos else 1
        ano_curso = st.selectbox("Ano do curso *", anos, index=ano_idx)
        turma = st.text_input("Turma *", value=pf.get("turma",""), placeholder="ex: T6")
    with col2:
        grupo = st.text_input("Grupo *", value=pf.get("grupo",""), placeholder="ex: Grupo A")
        try: data_def = datetime.date.fromisoformat(pf.get("data_inicio",""))
        except: data_def = datetime.date.today()
        data_inicio = st.date_input("Data de início (segunda-feira) *", value=data_def)
        num_semanas = st.number_input("Número de semanas *", 1, 20, int(pf.get("num_semanas",8)))

# BLOCO 2
with st.expander("👥 Bloco 2 — Alunos e Subgrupos", expanded=True):
    arquivo_alunos = st.file_uploader("Upload Excel de alunos (opcional)", type=["xlsx"])
    num_sg = st.number_input("Número de subgrupos", 2, 8, int(pf.get("num_sg",6)))
    alunos_por_sg = {}

    # Se importado, mostrar os alunos pré-preenchidos editáveis
    if pf.get("alunos_por_sg"):
        st.caption("✏️ Alunos importados — edite se necessário:")
        for sg, nomes in sorted(pf["alunos_por_sg"].items(), key=lambda x: int(x[0])):
            with st.expander(f"SG{sg} — {len(nomes)} alunos", expanded=False):
                txt = st.text_area(f"Alunos SG{sg}", value="\n".join(nomes), key=f"sg_imp_{sg}", height=100)
                alunos_por_sg[sg] = [n.strip() for n in txt.strip().split("\n") if n.strip()]

        # Opção de recarregar com outra opção de SGs
        if arquivo_alunos:
            col_sg1, col_sg2 = st.columns(2)
            with col_sg1:
                opcao_sg = st.selectbox("Mudar para:", ["Manter atual","4 Subgrupos","6 Subgrupos","8 Subgrupos"])
            with col_sg2:
                if opcao_sg != "Manter atual" and st.button("🔄 Recarregar da planilha"):
                    try:
                        n_alvo = int(opcao_sg.split()[0])
                        _bytes_alunos_imp = arquivo_alunos.read()
                        xls_a = pd.ExcelFile(io.BytesIO(_bytes_alunos_imp), engine="openpyxl")
                        sheets_g = [s for s in xls_a.sheet_names if "GRUPO" in s.upper()]
                        gp = pf.get("grupo","").upper().replace("GRUPO","").strip()
                        sh = next((s for s in sheets_g if gp in s.upper()), sheets_g[0] if sheets_g else None)
                        if sh:
                            df_r = pd.read_excel(io.BytesIO(_bytes_alunos_imp), engine="openpyxl", sheet_name=sh)
                            if "OPÇÃO" in df_r.columns:
                                op = next((o for o in df_r["OPÇÃO"].dropna().unique() if f"{n_alvo} SG" in str(o)), None)
                                if op:
                                    df_f = df_r[df_r["OPÇÃO"]==op]
                                    novos = {str(int(s)): df_f[df_f["Sub Grupo"]==s]["Nome Completo"].tolist()
                                             for s in sorted(df_f["Sub Grupo"].dropna().unique())}
                                    pf["alunos_por_sg"] = novos
                                    pf["num_sg"] = n_alvo
                                    st.session_state.prefill = pf
                                    st.success(f"✅ {n_alvo} SGs carregados!")
                                    st.rerun()
                    except Exception as e:
                        st.error(f"Erro: {e}")

    elif arquivo_alunos:
        try:
            _bytes_alunos2 = arquivo_alunos.read()
            xls2 = pd.ExcelFile(io.BytesIO(_bytes_alunos2), engine="openpyxl")
            grupos_disp = [s for s in xls2.sheet_names if "GRUPO" in s.upper()]
            grupo_sel = st.selectbox("Selecione o grupo", grupos_disp) if grupos_disp else None
            if grupo_sel:
                df_g = pd.read_excel(io.BytesIO(_bytes_alunos2), engine="openpyxl", sheet_name=grupo_sel)
                if "OPÇÃO" in df_g.columns:
                    opcoes = df_g["OPÇÃO"].dropna().unique()
                    opcao_sel = st.selectbox("Opção de subgrupos", opcoes)
                    df_f = df_g[df_g["OPÇÃO"] == opcao_sel]
                    for sg_num in sorted(df_f["Sub Grupo"].dropna().unique()):
                        nomes = df_f[df_f["Sub Grupo"]==sg_num]["Nome Completo"].tolist()
                        alunos_por_sg[str(int(sg_num))] = nomes
                    st.success(f"✅ {len(df_f)} alunos em {len(alunos_por_sg)} SGs")
        except Exception as e:
            st.error(f"Erro: {e}")

    if not alunos_por_sg:
        st.caption("Digite manualmente:")
        for sg in range(1, int(num_sg)+1):
            txt = st.text_area(f"SG{sg} (um nome por linha)", key=f"sg_{sg}", height=80)
            if txt.strip():
                alunos_por_sg[str(sg)] = [n.strip() for n in txt.strip().split("\n") if n.strip()]

# BLOCO 3 — Locais
with st.expander("📍 Bloco 3 — Blocos de Rodízio", expanded=True):
    pf_locais = pf.get("locais", [])
    num_locais_def = int(pf.get("num_locais", len(pf_locais) if pf_locais else 3))
    num_locais = st.number_input("Número de blocos de rodízio", 2, 8, num_locais_def,
        help="Cada bloco é um local de rodízio que pode conter 1 ou mais serviços vinculados.")
    locais = []

    # ── Configuração global do rodízio ──────────────────────────────────────
    st.markdown("---")
    st.markdown("**🔢 Distribuição de semanas entre os blocos (rodízio global)**")
    st.caption(
        f"Defina quantas semanas cada SG passa em cada bloco. "
        f"A soma deve ser igual ao total de semanas ({int(num_semanas)})."
    )

    # Recuperar ou inicializar distribuição salva
    key_rot = "rotacao_global"
    if key_rot not in st.session_state or len(st.session_state[key_rot]) != int(num_locais):
        # Distribuição padrão: divide igualmente, resto vai para os primeiros blocos
        base = int(num_semanas) // int(num_locais)
        resto = int(num_semanas) % int(num_locais)
        st.session_state[key_rot] = [base + (1 if j < resto else 0) for j in range(int(num_locais))]

    cols_rot = st.columns(int(num_locais))
    rotacao_semanas = []
    for j in range(int(num_locais)):
        nome_bloco_j = pf_locais[j].get("nome_bloco", pf_locais[j].get("nome", f"Bloco {j+1}")) if j < len(pf_locais) else f"Bloco {j+1}"
        with cols_rot[j]:
            val = st.number_input(
                f"Bloco {j+1}",
                min_value=1, max_value=int(num_semanas),
                value=st.session_state[key_rot][j],
                key=f"rot_sem_{j}",
                help=nome_bloco_j
            )
            rotacao_semanas.append(val)
            st.caption(nome_bloco_j[:20])

    soma_rot = sum(rotacao_semanas)
    if soma_rot == int(num_semanas):
        st.success(f"✅ {' + '.join(str(v) for v in rotacao_semanas)} = {soma_rot} semanas — OK!")
        st.session_state[key_rot] = rotacao_semanas
    elif soma_rot < int(num_semanas):
        st.warning(f"⚠️ Soma atual: {soma_rot} de {int(num_semanas)} semanas — faltam {int(num_semanas)-soma_rot}")
    else:
        st.error(f"❌ Soma atual: {soma_rot} — excede {int(num_semanas)} semanas em {soma_rot-int(num_semanas)}")

    st.markdown("---")

    def _servico_form(key_prefix, pl_srv, label, expanded=True):
        """Renderiza formulário de um serviço dentro de um bloco."""
        with st.expander(f"⚙️ {label}", expanded=expanded):
            ca, cb, cc = st.columns(3)
            with ca:
                nome = st.text_input("Nome", value=pl_srv.get("nome",""), key=f"{key_prefix}_nome", placeholder="ex: Enfermaria")
                abrev = st.text_input("Abreviação", value=pl_srv.get("abrev",""), key=f"{key_prefix}_abrev", placeholder="ex: Enf")
            with cb:
                obs = st.text_input("Observações", value=pl_srv.get("obs",""), key=f"{key_prefix}_obs")
                quem = st.text_input("Quem faz?", value=pl_srv.get("quem",""), key=f"{key_prefix}_quem",
                    placeholder="ex: todos | apenas SG par | alunos do Amb")
            with cc:
                st.caption("ℹ️ A duração por SG é configurada no nível do bloco, não do serviço individual.")
                duracao_sgs = pl_srv.get("duracao_por_sg", {})
                n_sgs_srv = pl_srv.get("n_sgs", len(alunos_por_sg) if alunos_por_sg else int(num_sg))

            st.markdown("**⏰ Dias Úteis:**")
            tu1, tu2, tu3 = st.columns(3)
            with tu1:
                st.markdown("**🌅 Manhã**")
                tem_m = st.checkbox("Tem manhã?", value=bool(pl_srv.get("manha","07-13h")), key=f"{key_prefix}_tm")
                if tem_m:
                    hor_m = st.text_input("Horário", value=pl_srv.get("manha","07-13h"), key=f"{key_prefix}_hm")
                    min_m = st.number_input("Mín/dia", 0, 20, int(pl_srv.get("min_manha",0)), key=f"{key_prefix}_mnm")
                    max_m = st.number_input("Máx/dia", 0, 20, int(pl_srv.get("max_manha",6)), key=f"{key_prefix}_mxm")
                    # Bloqueios manhã
                    key_bm = f"{key_prefix}_bloqs_m"
                    if key_bm not in st.session_state:
                        st.session_state[key_bm] = pl_srv.get("bloqueios_manha", [])
                    col_bm1, col_bm2 = st.columns(2)
                    with col_bm1:
                        if st.button("➕ Bloqueio manhã", key=f"{key_prefix}_add_bm"):
                            st.session_state[key_bm].append({"dia":"Qui","tipo":"Sem manhã","horario":""}); st.rerun()
                    with col_bm2:
                        if st.session_state[key_bm] and st.button("🗑️", key=f"{key_prefix}_clr_bm"):
                            st.session_state[key_bm] = []; st.rerun()
                    bloqueios_m = []
                    dias_op = ["Seg","Ter","Qua","Qui","Sex"]
                    for b, bloq in enumerate(st.session_state[key_bm]):
                        bm1,bm2,bm3,bm4 = st.columns([2,2,2,1])
                        with bm1: d_b = st.selectbox("Dia", dias_op, index=dias_op.index(bloq.get("dia","Qui")) if bloq.get("dia") in dias_op else 3, key=f"{key_prefix}_dbm_{b}")
                        with bm2: t_b = st.selectbox("Tipo", ["Sem manhã","Hor. reduzido"], index=1 if bloq.get("tipo")=="Hor. reduzido" else 0, key=f"{key_prefix}_tbm_{b}")
                        with bm3: h_b = st.text_input("Horário", value=bloq.get("horario",""), key=f"{key_prefix}_hbm_{b}") if t_b=="Hor. reduzido" else ""
                        with bm4:
                            st.write("")
                            if st.button("❌", key=f"{key_prefix}_delbm_{b}"): st.session_state[key_bm].pop(b); st.rerun()
                        bloqueios_m.append({"dia":d_b,"tipo":t_b,"horario":h_b})
                else:
                    hor_m,min_m,max_m,bloqueios_m = "",0,0,[]

            with tu2:
                st.markdown("**🌇 Tarde**")
                tem_t = st.checkbox("Tem tarde?", value=bool(pl_srv.get("tarde","12-18h")), key=f"{key_prefix}_tt")
                if tem_t:
                    hor_t = st.text_input("Horário normal", value=pl_srv.get("tarde","12-18h"), key=f"{key_prefix}_ht")
                    min_t = st.number_input("Mín/dia", 0, 20, int(pl_srv.get("min_tarde",3)), key=f"{key_prefix}_mnt")
                    max_t = st.number_input("Máx/dia", 0, 20, int(pl_srv.get("max_tarde",4)), key=f"{key_prefix}_mxt")
                    # Bloqueios tarde
                    key_bt = f"{key_prefix}_bloqs_t"
                    if key_bt not in st.session_state:
                        st.session_state[key_bt] = pl_srv.get("bloqueios_tarde", [{"dia":"Qui","tipo":"Sem tarde","horario":""}])
                    col_bt1, col_bt2 = st.columns(2)
                    with col_bt1:
                        if st.button("➕ Bloqueio tarde", key=f"{key_prefix}_add_bt"):
                            st.session_state[key_bt].append({"dia":"Seg","tipo":"Sem tarde","horario":""}); st.rerun()
                    with col_bt2:
                        if st.session_state[key_bt] and st.button("🗑️", key=f"{key_prefix}_clr_bt"):
                            st.session_state[key_bt] = []; st.rerun()
                    if not st.session_state[key_bt]: st.caption("Sem bloqueios.")
                    bloqueios_t = []
                    dias_op = ["Seg","Ter","Qua","Qui","Sex"]
                    for b, bloq in enumerate(st.session_state[key_bt]):
                        bt1,bt2,bt3,bt4 = st.columns([2,2,2,1])
                        with bt1: d_b = st.selectbox("Dia", dias_op, index=dias_op.index(bloq.get("dia","Qui")) if bloq.get("dia") in dias_op else 3, key=f"{key_prefix}_dbt_{b}")
                        with bt2: t_b = st.selectbox("Tipo", ["Sem tarde","Horário reduzido"], index=1 if bloq.get("tipo")=="Horário reduzido" else 0, key=f"{key_prefix}_tbt_{b}")
                        with bt3: h_b = st.text_input("Horário", value=bloq.get("horario","12-16h"), key=f"{key_prefix}_hbt_{b}") if t_b=="Horário reduzido" else ""
                        with bt4:
                            st.write("")
                            if st.button("❌", key=f"{key_prefix}_delbt_{b}"): st.session_state[key_bt].pop(b); st.rerun()
                        bloqueios_t.append({"dia":d_b,"tipo":t_b,"horario":h_b})
                else:
                    hor_t,min_t,max_t,bloqueios_t = "",0,0,[]

            with tu3:
                st.markdown("**🌙 Cinderela**")
                tem_c = st.checkbox("Tem cinderela?", value=bool(pl_srv.get("cinderela","")), key=f"{key_prefix}_tc")
                if tem_c:
                    hor_c = st.text_input("Horário", value=pl_srv.get("cinderela","19-23h"), key=f"{key_prefix}_hc")
                    min_c = st.number_input("Mín/dia", 0, 10, int(pl_srv.get("min_cind",0)), key=f"{key_prefix}_mnc")
                    max_c = st.number_input("Máx/dia", 0, 10, int(pl_srv.get("max_cind",2)), key=f"{key_prefix}_mxc")
                    dias_c_validos = ["Seg","Ter","Qua","Qui","Sex"]
                    dias_c_def = [d for d in pl_srv.get("dias_cind",["Sex"]) if d in dias_c_validos] or ["Sex"]
                    dias_c = st.multiselect("Dias", dias_c_validos, default=dias_c_def, key=f"{key_prefix}_dc")
                else:
                    hor_c,min_c,max_c,dias_c = "",0,0,[]

            st.markdown("**🏖️ Final de Semana:**")
            tf1, tf2, tf3 = st.columns(3)
            with tf1:
                st.markdown("**🌅 Manhã FDS**")
                tem_fm = st.checkbox("Tem?", value=bool(pl_srv.get("fds_manha","")), key=f"{key_prefix}_tfm")
                hor_fm = st.text_input("Horário", value=pl_srv.get("fds_manha","07-12h"), key=f"{key_prefix}_hfm") if tem_fm else ""
                min_fm = st.number_input("Mín", 0,10,int(pl_srv.get("fds_min_manha",1)), key=f"{key_prefix}_mnfm") if tem_fm else 0
                max_fm = st.number_input("Máx", 0,10,int(pl_srv.get("fds_max_manha",2)), key=f"{key_prefix}_mxfm") if tem_fm else 0
            with tf2:
                st.markdown("**🌇 Tarde FDS**")
                tem_ft = st.checkbox("Tem?", value=bool(pl_srv.get("fds_tarde","")), key=f"{key_prefix}_tft")
                hor_ft = st.text_input("Horário", value=pl_srv.get("fds_tarde","13-19h"), key=f"{key_prefix}_hft") if tem_ft else ""
                min_ft = st.number_input("Mín", 0,10,int(pl_srv.get("fds_min_tarde",1)), key=f"{key_prefix}_mnft") if tem_ft else 0
                max_ft = st.number_input("Máx", 0,10,int(pl_srv.get("fds_max_tarde",2)), key=f"{key_prefix}_mxft") if tem_ft else 0
            with tf3:
                st.markdown("**🌙 Cinderela FDS**")
                tem_fc = st.checkbox("Tem?", value=bool(pl_srv.get("fds_cind","")), key=f"{key_prefix}_tfc")
                hor_fc = st.text_input("Horário", value=pl_srv.get("fds_cind","19-23h"), key=f"{key_prefix}_hfc") if tem_fc else ""
                min_fc = st.number_input("Mín", 0,10,int(pl_srv.get("fds_min_cind",0)), key=f"{key_prefix}_mnfc") if tem_fc else 0
                max_fc = st.number_input("Máx", 0,10,int(pl_srv.get("fds_max_cind",2)), key=f"{key_prefix}_mxfc") if tem_fc else 0

            quem_fds = st.text_input("Quem faz FDS?", value=pl_srv.get("fds_quem",""), key=f"{key_prefix}_fdsquem")
            comp_fds = st.text_input("Compensação FDS?", value=pl_srv.get("fds_comp",""), key=f"{key_prefix}_fdscomp")

            return {
                "nome": nome, "abrev": abrev, "obs": obs, "quem": quem,
                "n_sgs": int(n_sgs_srv),
                "duracao_por_sg": duracao_sgs,
                "manha": hor_m, "min_manha": int(min_m), "max_manha": int(max_m), "bloqueios_manha": bloqueios_m,
                "tarde": hor_t, "min_tarde": int(min_t), "max_tarde": int(max_t), "bloqueios_tarde": bloqueios_t,
                "cinderela": hor_c, "min_cind": int(min_c), "max_cind": int(max_c), "dias_cind": dias_c,
                "fds": bool(tem_fm or tem_ft or tem_fc),
                "fds_manha": hor_fm, "fds_min_manha": int(min_fm), "fds_max_manha": int(max_fm),
                "fds_tarde": hor_ft, "fds_min_tarde": int(min_ft), "fds_max_tarde": int(max_ft),
                "fds_cind": hor_fc, "fds_min_cind": int(min_fc), "fds_max_cind": int(max_fc),
                "fds_quem": quem_fds, "fds_comp": comp_fds,
                "cov_tarde": int(min_t) if tem_t else 0,
            }

    for i in range(int(num_locais)):
        pl = pf_locais[i] if i < len(pf_locais) else {}
        with st.container():
            st.markdown(f"---")

            # Nome do bloco
            col_bloco_a, col_bloco_b = st.columns([2,3])
            with col_bloco_a:
                nome_bloco = st.text_input(
                    f"Nome do Bloco {i+1}",
                    value=pl.get("nome_bloco", pl.get("nome", f"Bloco {i+1}")),
                    key=f"nome_bloco_{i}",
                    placeholder=f"ex: BP | Ambulatório | SC Limeira"
                )
            with col_bloco_b:
                st.markdown(f"## 🏥 {nome_bloco or f'Bloco {i+1}'}")

            # Serviço principal
            n_srv_preview = 1 + st.session_state.get(f"n_srv_{i}", 0)
            label_s0 = f"Serviço 1 de {nome_bloco}" if nome_bloco else f"Serviço 1 do Bloco {i+1}"
            srv_principal = _servico_form(f"b{i}_s0", pl, label_s0, expanded=True)

            # Serviços adicionais
            key_n_srv = f"n_srv_{i}"
            if key_n_srv not in st.session_state:
                n_srv_def = len(pl.get("servicos_extras", [])) if pl.get("servicos_extras") else 0
                if pl.get("servico2"): n_srv_def = max(n_srv_def, 1)
                st.session_state[key_n_srv] = n_srv_def

            srv_extras = []
            for j in range(st.session_state[key_n_srv]):
                pl_extra = {}
                if j == 0 and pl.get("servico2"):
                    pl_extra = pl.get("servico2_cfg", {"nome": pl.get("servico2",""), "obs": pl.get("servico2_obs","")})
                elif j < len(pl.get("servicos_extras",[])):
                    pl_extra = pl["servicos_extras"][j]
                label_sj = f"Serviço {j+2} de {nome_bloco}" if nome_bloco else f"Serviço {j+2} do Bloco {i+1}"
                srv_extra = _servico_form(f"b{i}_s{j+1}", pl_extra, label_sj, expanded=False)
                srv_extras.append(srv_extra)
                if st.button(f"❌ Remover Serviço {j+2}", key=f"rm_srv_{i}_{j}"):
                    st.session_state[key_n_srv] -= 1; st.rerun()

            if st.button(f"➕ Adicionar serviço a {nome_bloco or f'Bloco {i+1}'}", key=f"add_srv_{i}"):
                st.session_state[key_n_srv] += 1; st.rerun()

            # Duração e distribuição de SGs no nível do BLOCO
            st.markdown("**📅 Distribuição de SGs neste bloco:**")
            n_sgs_total = len(alunos_por_sg) if alunos_por_sg else int(num_sg)
            n_srv_bloco = 1 + st.session_state.get(key_n_srv, 0)
            pl_dur_bloco = pl.get("duracao_por_sg", {})

            # Valor padrão vem do rodízio global configurado acima
            default_sem_sg = rotacao_semanas[i] if i < len(rotacao_semanas) else max(1, int(num_semanas) // max(int(num_locais), 1))

            col_dur1, col_dur2 = st.columns(2)
            with col_dur1:
                sem_por_sg = st.number_input(
                    f"Semanas que cada SG passa neste bloco",
                    min_value=1, max_value=int(num_semanas),
                    value=int(pl.get("sem_por_sg", default_sem_sg)),
                    key=f"sem_sg_{i}",
                    help=f"Quantas semanas cada SG fica neste bloco (passando por todos os serviços). Ex: {int(num_semanas)} sem ÷ {int(num_locais)} blocos = {default_sem_sg} sem/bloco por SG."
                )
            with col_dur2:
                # SGs presentes no bloco AO MESMO TEMPO (rodízio): total × semanas_no_bloco / total_semanas.
                # Ex.: 6 SGs, 2 sem neste bloco, 8 sem totais -> 6*2/8 = 1,5 -> ~2 SGs por vez (não os 6).
                sgs_simult = max(round(n_sgs_total * int(sem_por_sg) / max(int(num_semanas), 1)), 1)
                sgs_por_srv = max(sgs_simult // n_srv_bloco, 1) if n_srv_bloco > 0 else sgs_simult
                if n_srv_bloco > 1:
                    st.info(
                        f"📊 **{sgs_simult} SG(s) ao mesmo tempo** neste bloco (~{sgs_por_srv} por serviço) · "
                        f"cada SG fica **{sem_por_sg} sem** aqui · "
                        f"todos os {n_sgs_total} SGs passam por aqui ao longo do rodízio"
                    )
                else:
                    st.info(
                        f"📊 **{sgs_simult} SG(s) ao mesmo tempo** neste bloco · "
                        f"cada SG fica **{sem_por_sg} sem** aqui · "
                        f"todos os {n_sgs_total} SGs passam por aqui ao longo do rodízio"
                    )

            # Gerar duracao_sgs automaticamente
            duracao_sgs_bloco = {str(sg+1): int(sem_por_sg) for sg in range(n_sgs_total)}

            # Propagar para o serviço principal
            srv_principal["duracao_por_sg"] = duracao_sgs_bloco
            srv_principal["sem_por_sg"] = int(sem_por_sg)
            srv_principal["sgs_por_servico"] = sgs_por_srv
            for srv_e in srv_extras:
                srv_e["duracao_por_sg"] = duracao_sgs_bloco
                srv_e["sem_por_sg"] = int(sem_por_sg)

            # Montar dados do bloco
            bloco_nome = srv_principal["nome"]
            if srv_extras:
                bloco_nome = f"{srv_principal['nome']} + " + " + ".join([s["nome"] for s in srv_extras if s["nome"]])
            srv_principal["servicos_extras"] = srv_extras
            srv_principal["nome_bloco"] = nome_bloco
            bloco_nome = nome_bloco or srv_principal["nome"]
            locais.append(srv_principal)

# BLOCO 4 — Rodízio
with st.expander("🔄 Bloco 4 — Tabela de Rodízio", expanded=True):
    def gerar_sugestoes(n_sg, n_locais, n_sem, nomes_loc):
        locs = nomes_loc if nomes_loc and all(nomes_loc) else [f"Local{j+1}" for j in range(n_locais)]
        sugestoes = []
        if n_sg % n_locais == 0:
            tam = n_sg // n_locais
            sems_loc = n_sem // n_locais; resto = n_sem % n_locais
            linhas = []
            for p in range(n_locais):
                sgs = "+".join([f"SG{s}" for s in range(p*tam+1,(p+1)*tam+1)])
                sem_at = 1; rot = []
                for j in range(n_locais):
                    q = sems_loc + (1 if j < resto else 0)
                    rot.append(f"{locs[(p+j)%n_locais]} S{sem_at}-{sem_at+q-1}"); sem_at += q
                linhas.append(f"Par{p+1} = {sgs}: {' → '.join(rot)}")
            sugestoes.append({"nome": f"🔄 Por pares ({tam} SG/par)", "texto": "\n".join(linhas)})
        sems_loc2 = n_sem//n_locais; resto2 = n_sem%n_locais
        linhas2 = []
        for sg in range(1,n_sg+1):
            sem_at=1; rot=[]
            for j in range(n_locais):
                q=sems_loc2+(1 if j<resto2 else 0)
                rot.append(f"{locs[(sg-1+j)%n_locais]} S{sem_at}-{sem_at+q-1}"); sem_at+=q
            linhas2.append(f"SG{sg}: {' → '.join(rot)}")
        sugestoes.append({"nome":"🔄 Individual (cada SG passa por todos)","texto":"\n".join(linhas2)})
        return sugestoes

    nomes_loc_atual = [l.get("nome","") for l in locais]
    n_sg_atual = len(alunos_por_sg) if alunos_por_sg else int(num_sg)

    if st.button("💡 Sugerir rodízio automaticamente", use_container_width=True):
        st.session_state.sugestoes_rod = gerar_sugestoes(n_sg_atual, int(num_locais), int(num_semanas), nomes_loc_atual)

    if st.session_state.get("sugestoes_rod"):
        for idx, sug in enumerate(st.session_state.sugestoes_rod):
            with st.expander(sug["nome"], expanded=(idx==0)):
                st.code(sug["texto"])
                if st.button(f"✅ Usar esta", key=f"usar_{idx}"):
                    st.session_state.rodizio_escolhido = sug["texto"]; st.success("Aplicado!")

    rodizio_desc = st.text_area("Tabela de rodízio (edite à vontade)",
        value=st.session_state.get("rodizio_escolhido", pf.get("rodizio_desc","")),
        height=150, placeholder="Ex:\nPar1 = SG1+SG2: Enf S1-3 → PS S4-5 → Amb S6-8")

# BLOCO 5
with st.expander("⚙️ Bloco 5 — Regras Especiais", expanded=True):
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        regra_quinta = st.text_input("Quinta-feira", value=pf.get("regra_quinta","Sem tarde (ENAMED para todos)"))
        regra_terca = st.text_input("Terça-feira", value=pf.get("regra_terca","Tarde encurtada 12-16h (aula às 16h)"))
        limite_ch = st.number_input("Limite CH padrão (h)", 20, 60, int(pf.get("limite_ch",40)))
    with col_r2:
        limite_abs = st.number_input("Limite CH absoluto (h)", 20, 60, int(pf.get("limite_abs",43)))
        regra_fds = st.text_area("Regras de plantão FDS", value=pf.get("regra_fds",""), height=80)
        regras_extras = st.text_area("Outras regras", value=pf.get("regras_extras",""), height=80)

# BLOCO 6
with st.expander("📊 Bloco 6 — Formato do Excel", expanded=False):
    abas_excel = st.multiselect("Abas desejadas",
        ["Subgrupos","Calendário de Rodízio","Escala Nominal Detalhada","Resumo de Horas","Escala por Local","Regras e Restrições"],
        default=["Subgrupos","Calendário de Rodízio","Escala Nominal Detalhada","Resumo de Horas"])

# ── GERAR ────────────────────────────────────────────────────────────────────
st.divider()
if st.button("🚀 Gerar Escala com IA", type="primary", use_container_width=True):
    if not especialidade or not turma or not grupo:
        st.error("Preencha Especialidade, Turma e Grupo!")
    elif not alunos_por_sg:
        st.error("Adicione os alunos!")
    elif not rodizio_desc:
        st.error("Descreva o rodízio!")
    else:
        # Montar lista expandida de todos os serviços (incluindo vinculados)
        todos_servicos = []
        blocos_desc = []
        for idx_loc, loc in enumerate(locais):
            servs = [loc] + loc.get("servicos_extras", [])
            n_sgs_bloco = len(alunos_por_sg)
            nome_bl = loc.get("nome_bloco") or loc.get("nome","") or f"Bloco {idx_loc+1}"
            sem_sg_bloco = int(loc.get("sem_por_sg", max(1, int(num_semanas) // max(int(num_locais), 1))))
            # SGs presentes no bloco AO MESMO TEMPO (rodízio) — NÃO todos os SGs de uma vez.
            sgs_simult = max(round(n_sgs_bloco * sem_sg_bloco / max(int(num_semanas), 1)), 1)
            sgs_por_srv = max(sgs_simult // len(servs), 1) if servs else sgs_simult
            if len(servs) > 1:
                nomes = [s.get("nome","?") for s in servs]
                srv_desc_list = []
                for j, srv in enumerate(servs):
                    srv_desc_list.append(f"  - {srv.get('nome','?')} → ~{sgs_por_srv} SG(s) ao mesmo tempo")
                    srv_copy = dict(srv)
                    srv_copy["bloco"] = nome_bl
                    srv_copy["sgs_ao_mesmo_tempo"] = sgs_simult
                    todos_servicos.append(srv_copy)
                blocos_desc.append(
                    f"Bloco '{nome_bl}' ({' + '.join(nomes)}) — {len(servs)} serviços com RODÍZIO INTERNO:\n"
                    f"  • Em cada semana há ~{sgs_simult} SG(s) NESTE bloco ao mesmo tempo (NÃO os {n_sgs_bloco} SGs)\n"
                    f"  • Cada SG fica {sem_sg_bloco} semanas aqui; ao longo do rodízio TODOS os SGs passam por este bloco\n"
                    f"  • Os SG presentes se dividem entre os {len(servs)} serviços (~{sgs_por_srv} por serviço) e revezam\n" +
                    "\n".join(srv_desc_list) +
                    f"\n  ⚠️ NUNCA coloque o mesmo aluno em 2 serviços deste bloco no MESMO dia/turno"
                )
            else:
                srv_copy = dict(loc)
                srv_copy["bloco"] = nome_bl
                srv_copy["sgs_ao_mesmo_tempo"] = sgs_simult
                todos_servicos.append(srv_copy)
                blocos_desc.append(
                    f"Bloco '{nome_bl}': {loc.get('nome','?')} — serviço único · "
                    f"~{sgs_simult} SG(s) ao mesmo tempo (NÃO os {n_sgs_bloco} SGs); cada SG fica {sem_sg_bloco} sem aqui, "
                    f"e todos os SGs passam por aqui ao longo do rodízio"
                )

        briefing = f"""
# BRIEFING DE ESCALA MÉDICA

## IDENTIFICAÇÃO
Especialidade: {especialidade} | Ano: {ano_curso} | Turma: {turma} | Grupo: {grupo}
Início: {data_inicio} | Semanas: {num_semanas}

## ALUNOS POR SUBGRUPO
{json.dumps(alunos_por_sg, ensure_ascii=False, indent=2)}

## ESTRUTURA DE BLOCOS E SERVIÇOS
{chr(10).join(blocos_desc)}

## REGRA FUNDAMENTAL DE BLOCOS COM MÚLTIPLOS SERVIÇOS
Quando um bloco tem múltiplos serviços (ex: Enfermaria + PA Mandic):
- Os SGs se DIVIDEM simultaneamente entre os serviços (ex: SG1+SG2 na Enf enquanto SG3+SG4 no PA)
- Após algumas semanas, TROCAM de serviço dentro do bloco (rodízio interno)
- TODOS os SGs passarão por TODOS os serviços do bloco ao longo do tempo
- A exclusividade é TEMPORAL: no mesmo dia/turno, um aluno está em UM único serviço
- Na escala_detalhada, use o nome EXATO de cada serviço como "local"

## DETALHES DE CADA SERVIÇO
{json.dumps(todos_servicos, ensure_ascii=False, indent=2)}

## RODÍZIO
{rodizio_desc}

## REGRAS ESPECIAIS
Quinta: {regra_quinta}
Terça: {regra_terca}
Limite CH: {limite_ch}h/sem | Absoluto: {limite_abs}h
FDS: {regra_fds}
Extras: {regras_extras}
"""
        st.session_state.briefing_atual = briefing
        st.session_state.config_atual = {
            "especialidade": especialidade, "ano_curso": ano_curso,
            "grupo": grupo, "turma": turma,
            "data_inicio": str(data_inicio), "num_semanas": int(num_semanas),
            "locais": locais, "alunos_por_sg": alunos_por_sg,
            "rodizio_desc": rodizio_desc,
            "regras_especiais": {"quinta": regra_quinta, "terca": regra_terca,
                "limite_ch": int(limite_ch), "limite_abs": int(limite_abs), "fds": regra_fds},
            "pares": [], "blocos": [],
        }
        with st.spinner("Passo 1/2 — Calendário e resumo de horas... ⏳"):
            resposta = chamar_claude(
                [{"role": "user", "content": f"Gere a escala:\n{briefing}"}],
                system_prompt=SYSTEM_GERAR, max_tokens=8000
            )
        if resposta:
            dados1 = extrair_json(resposta) or {}
            st.session_state.esp_atual = especialidade
            st.session_state.grupo_atual = grupo
            st.session_state.turma_atual = turma
            cal_gerado = json.dumps(dados1.get("calendario_rodizio",[]), ensure_ascii=False)

            st.caption("Passo 2/2 — Escala detalhada (gerada semana a semana para não truncar)")
            det = gerar_detalhada_por_semana(briefing, cal_gerado, alunos_por_sg,
                                             st.session_state.config_atual, num_semanas)

            if det:
                dados1["escala_detalhada"] = det
                dados1["resumo_horas"] = recalcular_resumo_horas(dados1, st.session_state.config_atual)
                st.success(f"✅ Escala detalhada: {len(det)} entradas")

                # Validação real + 1 passe rápido de rebalanceamento (em paralelo).
                # Passes adicionais ficam a cargo do botão "Rebalancear" (você decide quando).
                val0 = validar_escala(dados1, st.session_state.config_atual)
                if not val0["ok"]:
                    n_est, n_con = len(val0["estouros"]), len(val0["conflitos"])
                    with st.spinner(f"⚖️ Rebalanceando turnos ({n_est} estouro(s) de CH, {n_con} conflito(s))... ⏳"):
                        dados1, _ = corrigir_escala_loop(dados1, st.session_state.config_atual, briefing, max_rodadas=1)
            else:
                st.warning("⚠️ Passo 2 não retornou dados. Tente o botão 'Gerar escala detalhada agora' abaixo.")

            st.session_state.escala_gerada = json.dumps(dados1, ensure_ascii=False)
            st.rerun()

# Mostrar resultado
if "escala_gerada" in st.session_state and "esp_atual" in st.session_state:
    st.divider()
    st.header("📊 Resultado")

    # Verificar se escala detalhada está vazia
    dados_atual = {}
    try:
        dados_atual = json.loads(st.session_state.escala_gerada) if isinstance(st.session_state.escala_gerada, str) else st.session_state.escala_gerada
    except: pass

    if not dados_atual.get("escala_detalhada"):
        st.warning("⚠️ Escala detalhada vazia — as abas Subgrupo e Individual não terão dados.")
        if st.button("🔄 Gerar escala detalhada agora", type="primary"):
            briefing_atual = st.session_state.get("briefing_atual","")
            cal = json.dumps(dados_atual.get("calendario_rodizio",[]), ensure_ascii=False)
            cfg_atual = st.session_state.get("config_atual",{})
            alunos_atual = cfg_atual.get("alunos_por_sg",{})
            n_sem_atual = cfg_atual.get("num_semanas",8)
            det = gerar_detalhada_por_semana(briefing_atual, cal, alunos_atual, cfg_atual, n_sem_atual)
            if det:
                dados_atual["escala_detalhada"] = det
                dados_atual["resumo_horas"] = recalcular_resumo_horas(dados_atual, cfg_atual)
                val0 = validar_escala(dados_atual, cfg_atual)
                if not val0["ok"]:
                    with st.spinner("⚖️ Rebalanceando turnos para respeitar as regras... ⏳"):
                        dados_atual, _ = corrigir_escala_loop(dados_atual, cfg_atual, briefing_atual, max_rodadas=1)
                st.session_state.escala_gerada = json.dumps(dados_atual, ensure_ascii=False)
                st.success(f"✅ {len(det)} entradas geradas!")
                st.rerun()
            else:
                st.error("A IA não retornou dados. Tente novamente.")

    mostrar_resultado(
        st.session_state.escala_gerada,
        st.session_state.get("esp_atual",""),
        st.session_state.get("grupo_atual",""),
        st.session_state.get("turma_atual","")
    )

# ════════════════════════════════════════════════════════════════════════════