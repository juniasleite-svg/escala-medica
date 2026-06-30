import streamlit as st
import pandas as pd
import json
import datetime
import requests
import csv
import io

st.set_page_config(page_title="Gerador de Escalas Médicas", page_icon="🏥", layout="wide")

# ── API ──────────────────────────────────────────────────────────────────────
def _claude_http(api_key, mensagens, system_prompt="", max_tokens=8000, tentativas=4):
    """Chamada crua à API — thread-safe (NÃO usa st.*). Re-tenta em rate limit / erro transitório."""
    if not api_key:
        return None
    import time
    for t in range(tentativas):
        try:
            resp = requests.post(
                "https://api.anthropic.com/v1/messages",
                headers={"x-api-key": api_key, "anthropic-version": "2023-06-01", "content-type": "application/json"},
                json={"model": "claude-sonnet-4-6", "max_tokens": max_tokens,
                      "system": system_prompt, "messages": mensagens},
                timeout=180
            )
            if resp.status_code in (429, 500, 502, 503, 529):  # rate limit / sobrecarga
                time.sleep(2 * (t + 1))
                continue
            result = resp.json()
        except Exception:
            time.sleep(2 * (t + 1))
            continue
        if "content" in result:
            return result["content"][0]["text"]
        time.sleep(2 * (t + 1))  # erro lógico (ex: overloaded no corpo) -> tenta de novo
    return None

def chamar_claude(mensagens, system_prompt="", max_tokens=8000):
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        st.error("Chave de API não configurada! Vá em Settings → Secrets.")
        return None
    texto = _claude_http(api_key, mensagens, system_prompt, max_tokens)
    if texto is None:
        st.error("Erro API: a IA não respondeu (timeout ou resposta inválida).")
    return texto

def _ler_subgrupos(df_f):
    """Lê um recorte da planilha de alunos -> ({sg: [nomes]}, {nome: RA})."""
    apsg, ra = {}, {}
    for sg_num in sorted(df_f["Sub Grupo"].dropna().unique(), key=lambda x: int(x)):
        sub = df_f[df_f["Sub Grupo"] == sg_num]
        nomes = [str(x).strip() for x in sub["Nome Completo"].tolist() if str(x).strip()]
        apsg[str(int(sg_num))] = nomes
        if "RA" in sub.columns:
            for _, rr in sub.iterrows():
                nm = str(rr["Nome Completo"]).strip()
                if nm:
                    ra[nm] = str(rr.get("RA", "") or "").strip()
    return apsg, ra

# ── Formato "tabela longa" de alunos ──────────────────────────────────────────
# Algumas bases trazem UMA aba por OPÇÃO de subgrupos (ex.: "Opção 3 (8 SG)"),
# com TODOS os grupos juntos numa coluna "Grupo" e todos os rodízios numa coluna
# "Nº Rodízio". A divisão de subgrupos é constante entre rodízios, então basta
# filtrar pelo grupo e por um único rodízio. (O outro formato — uma aba por grupo,
# com coluna "OPÇÃO" — continua suportado pelos chamadores.)
def _abas_long_alunos(xls):
    """Abas no formato tabela longa (colunas Grupo + Sub Grupo + Nome Completo)."""
    out = []
    for s in xls.sheet_names:
        try:
            hd = xls.parse(s, nrows=0)
        except Exception:
            continue
        cols = {str(c).strip().lower() for c in hd.columns}
        if "grupo" in cols and "sub grupo" in cols and "nome completo" in cols:
            out.append(s)
    return out

def _df_long_grupo(xls_bytes, sheet, grupo):
    """Recorta a tabela longa: só o grupo escolhido e um único rodízio (divisão constante)."""
    d = pd.read_excel(io.BytesIO(xls_bytes), engine="openpyxl", sheet_name=sheet)
    d.columns = [str(c).strip() for c in d.columns]
    if "Grupo" in d.columns and grupo:
        d = d[d["Grupo"].astype(str).str.strip() == str(grupo).strip()]
    if "Nº Rodízio" in d.columns and d["Nº Rodízio"].notna().any():
        d = d[d["Nº Rodízio"] == d["Nº Rodízio"].dropna().iloc[0]]
    if "Nome Completo" in d.columns:
        d = d.drop_duplicates(subset=["Nome Completo"])
    return d

def _grupos_long(xls_bytes, sheet):
    """Lista de grupos (ex.: 'Grupo A'… 'Grupo F') presentes numa aba long."""
    d = pd.read_excel(io.BytesIO(xls_bytes), engine="openpyxl", sheet_name=sheet)
    d.columns = [str(c).strip() for c in d.columns]
    if "Grupo" not in d.columns:
        return []
    vals = [str(x).strip() for x in d["Grupo"].dropna().unique() if str(x).strip()]
    return sorted(set(vals))

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
    partes = str(e.get("horario", "")).split("-")
    if len(partes) >= 2:
        a = _hhmm_para_horas(partes[0])
        b = _hhmm_para_horas(partes[1])
        if a is not None and b is not None and b > a:
            return float(b - a)
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

def _clamp(v, lo, hi, default=0):
    """Garante que o valor caiba em [lo, hi] (evita StreamlitValueAboveMaxError)."""
    try:
        v = int(v)
    except (TypeError, ValueError):
        v = int(default)
    return max(lo, min(v, hi))

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
            def _faixa(mn, mx):
                try: mn = int(mn or 0)
                except (TypeError, ValueError): mn = 0
                try: mx = int(mx) if mx not in (None, "", 0) else None
                except (TypeError, ValueError): mx = None
                return (mn, mx)
            turnos = {}      # dias úteis
            for tk, (mn, mx, ativo) in {
                "manha": (s.get("min_manha"), s.get("max_manha"), s.get("manha")),
                "tarde": (s.get("min_tarde"), s.get("max_tarde"), s.get("tarde")),
                "cind":  (s.get("min_cind"),  s.get("max_cind"),  s.get("cinderela")),
            }.items():
                if ativo:
                    turnos[tk] = _faixa(mn, mx)
            turnos_fds = {}  # fim de semana (mín/máx próprios)
            for tk, (mn, mx, ativo) in {
                "manha": (s.get("fds_min_manha"), s.get("fds_max_manha"), s.get("fds_manha")),
                "tarde": (s.get("fds_min_tarde"), s.get("fds_max_tarde"), s.get("fds_tarde")),
                "cind":  (s.get("fds_min_cind"),  s.get("fds_max_cind"),  s.get("fds_cind")),
            }.items():
                if ativo:
                    turnos_fds[tk] = _faixa(mn, mx)
            servs.append({"nomes": nomes, "bloco": bloco,
                          "label": s.get("nome") or s.get("abrev") or "",
                          "turnos": turnos, "turnos_fds": turnos_fds})
    return servs

def validar_escala(dados, config):
    """Confere a escala de verdade: horas/semana, conflitos, cobertura, mín/máx por serviço e rodízio dentro do bloco."""
    det = dados.get("escala_detalhada") or []
    reg = config.get("regras_especiais", {})
    limite = int(reg.get("limite_ch", 40))
    limite_abs = int(reg.get("limite_abs", 43))
    alvo_min = int(reg.get("limite_min", 34))
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
    subcarga = []
    for (al, sem), h in horas.items():
        if h > limite_abs:
            estouros.append({"aluno": al, "semana": sem, "horas": round(h, 1), "limite": limite_abs, "nivel": "absoluto"})
        elif h > limite:
            estouros.append({"aluno": al, "semana": sem, "horas": round(h, 1), "limite": limite, "nivel": "padrao"})
        elif h < alvo_min:
            subcarga.append({"aluno": al, "semana": sem, "horas": round(h, 1), "alvo": alvo_min})
    estouros.sort(key=lambda x: -x["horas"])
    subcarga.sort(key=lambda x: x["horas"])

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
        eh_fds = dia in ("Sáb", "Sab", "Dom")
        chave = "turnos_fds" if eh_fds else "turnos"
        srv = next((s for s in servs_cfg if turno in s.get(chave, {}) and _nome_bate(s["nomes"], _norm(local))), None)
        if not srv:
            continue
        mn, mx = srv[chave][turno]
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

    # ── Tempo por serviço por aluno (horas, % e quem não passou) ─────────────
    visitou = _dd(set)
    todos_locais = set()
    hser = _dd(lambda: _dd(float))  # aluno -> local -> horas
    htot = _dd(float)               # aluno -> horas totais
    for e in det:
        loc = str(e.get("local", ""))
        if not loc:
            continue
        todos_locais.add(loc)
        h = _horas_entrada(e)
        for al in _alunos_entrada(e):
            visitou[al].add(loc)
            hser[al][loc] += h
            htot[al] += h
    locais_lista = sorted(todos_locais)
    nao_passou = []
    for al in (visitou or {}):
        faltam = todos_locais - visitou[al]
        if faltam:
            nao_passou.append({"aluno": al, "faltam": sorted(faltam)})
    nao_passou.sort(key=lambda x: x["aluno"])
    pct_servico = {}
    for al in htot:
        tot = htot[al] or 1
        pct_servico[al] = {loc: round(hser[al].get(loc, 0) / tot * 100) for loc in locais_lista}

    # ── Sobrecarga de bloco numa semana: alunos demais p/ a capacidade → CH não fecha ──
    # (ex.: 4 SGs no mesmo Ambulatório na mesma semana → cada um só consegue ~18h)
    def _cap_bloco(loc):
        cap = 0.0
        d_tarde = 5 - (1 if "sem tarde" in regra_quinta else 0)
        for s in [loc] + (loc.get("servicos_extras") or []):
            for tk, hor, mx, dias in [
                ("manha", s.get("manha"), s.get("max_manha"), 5),
                ("tarde", s.get("tarde"), s.get("max_tarde"), d_tarde),
                ("cind", s.get("cinderela"), s.get("max_cind"), len(s.get("dias_cind") or []) or 5),
                ("manha", s.get("fds_manha"), s.get("fds_max_manha"), 2),
                ("tarde", s.get("fds_tarde"), s.get("fds_max_tarde"), 2),
                ("cind", s.get("fds_cind"), s.get("fds_max_cind"), 2),
            ]:
                if not hor:
                    continue
                try:
                    q = int(mx) if str(mx) not in ("None", "", "0") else 99
                except (TypeError, ValueError):
                    q = 99
                cap += dias * max(q, 1) * _dur_horario(hor, tk)
        return cap

    blocos_cfg = config.get("locais", [])
    caps_bloco = [_cap_bloco(lc) for lc in blocos_cfg]
    rotulos_bloco = []
    for lc in blocos_cfg:
        rot = {_norm(lc.get("nome_bloco")), _norm(lc.get("nome")), _norm(lc.get("abrev"))} - {""}
        for s in [lc] + (lc.get("servicos_extras") or []):
            rot |= ({_norm(s.get("nome")), _norm(s.get("abrev"))} - {""})
        rotulos_bloco.append(rot)
    alunos_bs = _dd(set)  # (bidx, semana) -> alunos distintos
    for e in det:
        if e.get("plantao"):      # plantões de complemento não contam como "cohorte" do bloco
            continue
        loc = _norm(e.get("local", ""))
        sem = e.get("semana", "?")
        bidx = next((bi for bi, rot in enumerate(rotulos_bloco)
                     if any(r and (r in loc or loc in r) for r in rot)), None)
        if bidx is None:
            continue
        for al in _alunos_entrada(e):
            alunos_bs[(bidx, sem)].add(al)
    plantoes_complemento = sum(1 for e in det if e.get("plantao"))
    sobrecarga_bloco = []
    for (bidx, sem), als in alunos_bs.items():
        n = len(als)
        if n <= 0:
            continue
        ch_max = caps_bloco[bidx] / n
        if ch_max < alvo_min - 0.5:
            sobrecarga_bloco.append({
                "bloco": blocos_cfg[bidx].get("nome_bloco") or blocos_cfg[bidx].get("nome") or f"Bloco {bidx+1}",
                "semana": sem, "alunos": n, "ch_max": round(ch_max, 1)})
    sobrecarga_bloco.sort(key=lambda x: (x["ch_max"], str(x["semana"])))

    semanas_ruins = sorted(
        {e["semana"] for e in estouros} | {c["semana"] for c in conflitos}
        | {d["semana"] for d in desvios} | {a["semana"] for a in ausentes},
        key=lambda x: int(x) if str(x).isdigit() else 999)
    return {
        "estouros": estouros, "conflitos": conflitos, "buracos": buracos,
        "desvios": desvios, "ausentes": ausentes, "subcarga": subcarga,
        "nao_passou": nao_passou, "pct_servico": pct_servico, "locais_lista": locais_lista,
        "sobrecarga_bloco": sobrecarga_bloco, "plantoes_complemento": plantoes_complemento,
        "limite": limite, "limite_abs": limite_abs, "alvo_min": alvo_min,
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
    ra_map = config.get("ra_por_aluno", {}) or {}
    nomes_todos = list(aluno_sg.keys()) or list(hs.keys())
    linhas = []
    for al in nomes_todos:
        semanas = [round(hs[al].get(i + 1, 0.0), 1) for i in range(num_sem)]
        base = dict(rh_old.get(al, {}))
        base.update({
            "sg": aluno_sg.get(al, "") or base.get("sg", ""),
            "nome": al, "ra": ra_map.get(al) or base.get("ra", ""),
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

def gerar_detalhada_por_semana(briefing, calendario_json, alunos_por_sg, config, num_semanas, ui=True, apenas_semanas=None):
    """Gera a escala_detalhada por semana, em paralelo, com re-tentativa. apenas_semanas: gera só essas."""
    import concurrent.futures
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    limite = config.get("regras_especiais", {}).get("limite_ch", 40)
    n = max(int(num_semanas), 1)
    alvo = sorted(apenas_semanas) if apenas_semanas else list(range(1, n + 1))
    total = len(alvo) or 1

    # Especificação RÍGIDA dos turnos válidos de cada serviço (impede a IA de inventar turno/FDS)
    def _spec_servicos():
        linhas = []
        for loc in config.get("locais", []):
            bloco = loc.get("nome_bloco") or loc.get("nome") or ""
            for s in [loc] + (loc.get("servicos_extras") or []):
                nome = s.get("nome") or s.get("abrev") or "?"
                parts = []
                if s.get("manha"):
                    parts.append(f"Manhã {s.get('manha')} (min {s.get('min_manha',0)}/máx {s.get('max_manha','-')} alunos)")
                if s.get("tarde"):
                    parts.append(f"Tarde {s.get('tarde')} (min {s.get('min_tarde',0)}/máx {s.get('max_tarde','-')} alunos)")
                if s.get("cinderela"):
                    dc = s.get("dias_cind") or []
                    parts.append(f"Cinderela {s.get('cinderela')}" + (f" SÓ em {'/'.join(dc)}" if dc else ""))
                nao = [t for t, k in [("manhã", "manha"), ("tarde", "tarde"), ("cinderela", "cinderela")] if not s.get(k)]
                tem_fds = bool(s.get("fds_manha") or s.get("fds_tarde") or s.get("fds_cind"))
                bloq = []
                for b in (s.get("bloqueios_manha") or []):
                    bloq.append(f"manhã {b.get('dia')}={b.get('tipo')}")
                for b in (s.get("bloqueios_tarde") or []):
                    bloq.append(f"tarde {b.get('dia')}={b.get('tipo')}")
                linha = f"- {nome} (bloco {bloco}): " + ("; ".join(parts) if parts else "sem turnos")
                if nao:
                    linha += f" | NÃO TEM: {', '.join(nao)}"
                linha += " | COM plantão de fim de semana" if tem_fds else " | SEM fim de semana (só Seg–Sex)"
                if bloq:
                    linha += f" | bloqueios: {', '.join(bloq)}"
                linhas.append(linha)
        return "\n".join(linhas)
    spec = _spec_servicos()

    def _msg(sem):
        return (
            f"Briefing:\n{briefing}\n\n"
            f"Calendário de rodízio (todas as semanas):\n{calendario_json}\n\n"
            f"Alunos por subgrupo:\n{json.dumps(alunos_por_sg, ensure_ascii=False)}\n\n"
            f"⛔ TURNOS VÁLIDOS DE CADA SERVIÇO — use SOMENTE estes. É PROIBIDO criar turno que o serviço "
            f"não tem (ex.: tarde numa enfermaria só-manhã) e PROIBIDO criar Sábado/Domingo em serviço "
            f"marcado 'SEM fim de semana':\n{spec}\n\n"
            f"Gere a escala_detalhada APENAS da SEMANA {sem}. Para CADA serviço ATIVO nesta semana "
            f"(veja o calendário), gere CADA turno VÁLIDO em TODOS os dias úteis Seg–Sex (exceto bloqueios), "
            f"com os alunos REVEZADOS para ninguém passar de {limite}h e respeitando o mín/máx por turno. "
            f"Divida os alunos entre os serviços de cada bloco. NÃO gere outras semanas."
        )

    resultados = {}
    barra = st.progress(0.0, text="Gerando escala detalhada (semanas em paralelo)...") if ui else None
    # Concorrência moderada (evita rate limit) + re-tentativa das semanas que vierem vazias.
    for rodada in range(3):
        pendentes = [s for s in alvo if s not in resultados]
        if not pendentes:
            break
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(4, len(pendentes))) as ex:
            futuros = {ex.submit(_claude_http, api_key, [{"role": "user", "content": _msg(sem)}],
                                 SYSTEM_DETALHE, 8000): sem for sem in pendentes}
            for fut in concurrent.futures.as_completed(futuros):
                sem = futuros[fut]
                novas = (extrair_json(fut.result()) or {}).get("escala_detalhada")
                if novas:
                    for e in novas:
                        e["semana"] = sem
                    resultados[sem] = novas
                if barra:
                    extra = "" if rodada == 0 else " (re-tentando)"
                    barra.progress(len(resultados) / total, text=f"Escala detalhada: {len(resultados)}/{total} semanas prontas{extra}")
    if barra:
        barra.empty()
    faltando = [s for s in alvo if s not in resultados]
    if faltando and ui:
        st.warning(f"⚠️ Não consegui gerar a(s) semana(s): {', '.join(map(str, faltando))}. "
                   f"Clique em 'Completar semanas faltantes' para tentar de novo.")
    todas = []
    for sem in alvo:
        todas.extend(resultados.get(sem, []))
    return todas

# ── Gerador DETERMINÍSTICO da escala detalhada (Python, sem IA) ───────────────
def _hhmm_para_horas(p):
    """Converte um lado do horário ('07', '07:30', '7h', '13h00') em horas decimais."""
    m = _re_val.findall(r"\d{1,2}", str(p or ""))
    if not m:
        return None
    h = int(m[0])
    mn = int(m[1]) if len(m) > 1 else 0
    return h + mn / 60.0

def _dur_horario(horario, turno_key=""):
    # aceita '07-13h', '07:00-12:00', '07:30-12:00', '13h-19h' etc. (lê hora E minutos)
    partes = str(horario or "").split("-")
    if len(partes) >= 2:
        a = _hhmm_para_horas(partes[0])
        b = _hhmm_para_horas(partes[1])
        if a is not None and b is not None and b > a:
            return float(b - a)
    return {"manha": 6.0, "tarde": 6.0, "cind": 4.0}.get(turno_key, 6.0)

def previa_viabilidade(locais, total_alunos, num_blocos, alvo_min, limite, regra_quinta=""):
    """ANTES de gerar: estima a CH/aluno alcançável por bloco (mín forçado e máx)."""
    rq = str(regra_quinta or "").lower()
    nb = max(int(num_blocos), 1)
    alunos_bloco = max(round(total_alunos / nb), 1)  # ~alunos simultâneos por bloco
    out = []
    for loc in locais:
        nome_bl = loc.get("nome_bloco") or loc.get("nome") or "Bloco"
        cap_min = cap_max = 0.0
        for s in [loc] + (loc.get("servicos_extras") or []):
            def add(tk, hor, mn, mx, dias):
                nonlocal cap_min, cap_max
                if not hor:
                    return
                h = _dur_horario(hor, tk)
                qmin = max(int(mn or 0), 1)
                try:
                    qmax = int(mx) if str(mx) not in ("None", "", "0") else alunos_bloco
                except (TypeError, ValueError):
                    qmax = alunos_bloco
                cap_min += dias * qmin * h
                cap_max += dias * max(qmax, qmin) * h
            d_tarde = 5 - (1 if "sem tarde" in rq else 0)
            add("manha", s.get("manha"), s.get("min_manha"), s.get("max_manha"), 5)
            add("tarde", s.get("tarde"), s.get("min_tarde"), s.get("max_tarde"), d_tarde)
            if s.get("cinderela"):
                add("cind", s.get("cinderela"), s.get("min_cind"), s.get("max_cind"),
                    len(s.get("dias_cind") or []) or 5)
            add("manha", s.get("fds_manha"), s.get("fds_min_manha"), s.get("fds_max_manha"), 2)
            add("tarde", s.get("fds_tarde"), s.get("fds_min_tarde"), s.get("fds_max_tarde"), 2)
            add("cind", s.get("fds_cind"), s.get("fds_min_cind"), s.get("fds_max_cind"), 2)
        out.append({"bloco": nome_bl, "alunos_bloco": alunos_bloco,
                    "ch_min": round(cap_min / alunos_bloco, 1),
                    "ch_max": round(cap_max / alunos_bloco, 1)})
    return out

def gerar_detalhada_python(calendario, config):
    """Monta a escala_detalhada dia a dia a partir do calendário + regras dos serviços.
    Garante: só turnos válidos, cobertura completa, exclusividade e equilíbrio de horas (≤ limite)."""
    from datetime import date, timedelta
    reg = config.get("regras_especiais", {})
    limite = int(reg.get("limite_ch", 40))
    limite_abs = int(reg.get("limite_abs", 43))
    alvo_min = int(reg.get("limite_min", 34))   # CH mínima alvo (completa até aqui se possível)
    regra_quinta = str(reg.get("quinta", "")).lower()
    alunos_por_sg = config.get("alunos_por_sg", {})
    num_sem = int(config.get("num_semanas", 8))
    try:
        d0 = date.fromisoformat(str(config.get("data_inicio", "")))
    except Exception:
        d0 = date.today()
    dias3 = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]

    # Blocos e seus serviços
    blocos = []
    for loc in config.get("locais", []):
        blocos.append({"nome": loc.get("nome_bloco") or loc.get("nome") or "",
                       "loc": loc, "servs": [loc] + (loc.get("servicos_extras") or [])})

    def _match(dest):
        a = _norm(dest)
        for b in blocos:
            nb = {_norm(b["nome"]), _norm(b["loc"].get("nome")), _norm(b["loc"].get("abrev"))} - {""}
            svc = None
            for s in b["servs"]:
                sn = {_norm(s.get("nome")), _norm(s.get("abrev"))} - {""}
                if any(x and x in a for x in sn):
                    svc = s
            if svc is not None or any(x and x in a for x in nb):
                return b, svc
        return None, None

    cal_by_week = {}
    for row in (calendario or []):
        try:
            cal_by_week[int(row.get("semana"))] = row.get("alocacao", {}) or {}
        except (TypeError, ValueError):
            continue

    def _slots_servico(s, datas):
        """Lista de (data, dia3, turno_nome, turno_key, horario, horas, qmin, qmax, local) do serviço."""
        bloq = set()
        reduz = {}   # (turno, dia3) -> horário reduzido naquele dia (ex.: terça 13-16h)
        for bm in (s.get("bloqueios_manha") or []):
            tp = str(bm.get("tipo", "")).lower()
            if tp.startswith("sem"):
                bloq.add(("manha", _dia_curto(bm.get("dia"))))
            elif "reduz" in tp and bm.get("horario"):
                reduz[("manha", _dia_curto(bm.get("dia")))] = bm.get("horario")
        for bt in (s.get("bloqueios_tarde") or []):
            tp = str(bt.get("tipo", "")).lower()
            if tp.startswith("sem"):
                bloq.add(("tarde", _dia_curto(bt.get("dia"))))
            elif "reduz" in tp and bt.get("horario"):
                reduz[("tarde", _dia_curto(bt.get("dia")))] = bt.get("horario")

        def _cnt(mn, mx):
            qmin = max(int(mn) if mn else 0, 1)   # todo turno presente cobre pelo menos 1
            try:
                qmax = int(mx) if mx not in (None, "", 0) else None
            except (TypeError, ValueError):
                qmax = None
            if qmax is not None and qmax < qmin:
                qmax = qmin
            return qmin, qmax

        locn = s.get("nome") or s.get("abrev") or ""
        uteis = []
        if s.get("manha"):
            qn, qx = _cnt(int(s.get("min_manha") or 0), s.get("max_manha"))
            uteis.append(("Manhã", "manha", s.get("manha"), _dur_horario(s.get("manha"), "manha"), qn, qx, None))
        if s.get("tarde"):
            qn, qx = _cnt(int(s.get("min_tarde") or 0), s.get("max_tarde"))
            uteis.append(("Tarde", "tarde", s.get("tarde"), _dur_horario(s.get("tarde"), "tarde"), qn, qx, None))
        if s.get("cinderela"):
            dc = {_dia_curto(x) for x in (s.get("dias_cind") or [])}
            qn, qx = _cnt(int(s.get("min_cind") or 0), s.get("max_cind"))
            uteis.append(("Cinderela", "cind", s.get("cinderela"), _dur_horario(s.get("cinderela"), "cind"), qn, qx, dc))
        fds = []
        if s.get("fds_manha"):
            qn, qx = _cnt(int(s.get("fds_min_manha") or 0), s.get("fds_max_manha"))
            fds.append(("Manhã", "manha", s.get("fds_manha"), _dur_horario(s.get("fds_manha"), "manha"), qn, qx, None))
        if s.get("fds_tarde"):
            qn, qx = _cnt(int(s.get("fds_min_tarde") or 0), s.get("fds_max_tarde"))
            fds.append(("Tarde", "tarde", s.get("fds_tarde"), _dur_horario(s.get("fds_tarde"), "tarde"), qn, qx, None))
        if s.get("fds_cind"):
            qn, qx = _cnt(int(s.get("fds_min_cind") or 0), s.get("fds_max_cind"))
            qn = int(s.get("fds_min_cind") or 0)   # cinderela de FDS: respeita o mín REAL (0) — prioriza manhã/tarde
            fds.append(("Cinderela", "cind", s.get("fds_cind"), _dur_horario(s.get("fds_cind"), "cind"), qn, qx, None))

        slots = []
        for di in range(7):
            dia3 = dias3[di]
            grupo = uteis if di < 5 else fds
            for (tn, tk, hor, hrs, qmin, qmax, dc) in grupo:
                if tk == "cind" and dc is not None and dia3 not in dc:
                    continue
                if (tk, dia3) in bloq:
                    continue
                if tk == "tarde" and dia3 == "Qui" and "sem tarde" in regra_quinta:
                    continue
                # horário reduzido naquele dia (ex.: terça à tarde até 16h) → ajusta horário e horas
                h_dia, hrs_dia = hor, hrs
                if (tk, dia3) in reduz:
                    h_dia = reduz[(tk, dia3)]
                    hrs_dia = _dur_horario(h_dia, tk)
                slots.append((datas[di], dia3, tn, tk, h_dia, hrs_dia, qmin, qmax, locn))
        return slots

    detalhada = []
    g_tcount = {}  # contagem GLOBAL por aluno e tipo de turno (equilíbrio ao longo do semestre)
    sg_de_aluno = {}
    for _sg, _nomes_sg in alunos_por_sg.items():
        for _n in _nomes_sg:
            sg_de_aluno[_n] = _sg
    for w in range(1, num_sem + 1):
        base = d0 + timedelta(weeks=w - 1)
        datas = [base + timedelta(days=i) for i in range(7)]
        aloc = cal_by_week.get(w, {})

        # SGs -> BLOCO nesta semana (o bloco é a unidade: alunos cobrem TODOS os serviços dele)
        bloco_sgs = {}  # bidx -> set(sg)
        for sg_key, dest in aloc.items():
            sgn = _re_val.sub(r"\D", "", str(sg_key))
            if not sgn:
                continue
            b, _svc = _match(dest)
            if not b:
                continue
            bloco_sgs.setdefault(blocos.index(b), set()).add(sgn)

        for bidx, sgset in bloco_sgs.items():
            b = blocos[bidx]
            estud = [(n, sg) for sg in sorted(sgset) for n in alunos_por_sg.get(sg, [])]
            if not estud:
                continue
            nomes = [n for n, _ in estud]
            sg_de = {n: sg for n, sg in estud}
            servs_bloco = b["servs"]
            # Todos os turnos de TODOS os serviços do bloco (guardando o serviço de cada slot)
            slots = []
            slot_svc = []
            for si, s in enumerate(servs_bloco):
                novos = _slots_servico(s, datas)
                slots.extend(novos)
                slot_svc.extend([si] * len(novos))

            # Plantão EXCLUSIVO: serviço cujo turno só pode ser feito por quem, na mesma semana,
            # está escalado em OUTRO serviço (de origem) do mesmo bloco. Mapeia si -> si_origem.
            _excl_por_si = {}
            _excl_turnos_por_si = {}   # si -> set de categorias (du_/fds_ × manha/tarde/cind); vazio = todas
            for _si, _s in enumerate(servs_bloco):
                _alvo = _norm(str(_s.get("exclusivo_de_servico", "") or ""))
                if not _alvo:
                    continue
                for _sj, _s2 in enumerate(servs_bloco):
                    if _sj == _si:
                        continue
                    _rot = {_norm(_s2.get("nome")), _norm(_s2.get("abrev"))} - {""}
                    if _alvo in _rot or any(r and (_alvo in r or r in _alvo) for r in _rot):
                        _excl_por_si[_si] = _sj
                        _excl_turnos_por_si[_si] = set(_s.get("exclusivo_turnos") or [])
                        break

            # Plano de continuidade POR SERVIÇO: a continuidade prende os MESMOS alunos
            # ao SERVIÇO escolhido (ex: Enfermaria) por N dias úteis seguidos; depois gira para
            # que todos passem por ele. IMPORTANTE: o aluno preso à Enfermaria de manhã CONTINUA
            # livre para ser escalado no PA à tarde / cinderela nos mesmos dias (a trava vale só
            # para os turnos do PRÓPRIO serviço de continuidade, não bloqueia os outros serviços).
            # Configurada POR BLOCO + SERVIÇO (Bloco 5), não global.
            dias_consec = int(b["loc"].get("dias_consec", 0) or 0)
            cont_nome = _norm(str(b["loc"].get("consec_servico", "") or ""))
            pinned = {}  # wd (0..4) -> set de alunos presos ao serviço de continuidade
            usa_continuidade = dias_consec > 0 and len(servs_bloco) > 1
            cont_idx = 0
            if usa_continuidade:
                # qual serviço do bloco recebe a continuidade (default = principal/serviço 1)
                if cont_nome:
                    for si, s in enumerate(servs_bloco):
                        rotulos = {_norm(s.get("nome")), _norm(s.get("abrev"))} - {""}
                        if cont_nome in rotulos or any(cont_nome in r or r in cont_nome for r in rotulos):
                            cont_idx = si; break
                # capacidade do serviço de continuidade (~quantos cabem por vez)
                s_cont = servs_bloco[cont_idx]
                cap_cont = (s_cont.get("max_manha") or s_cont.get("max_tarde")
                            or s_cont.get("max_cind") or s_cont.get("min_manha") or 1)
                try: cap_cont = int(cap_cont)
                except (TypeError, ValueError): cap_cont = 1
                cap_cont = max(1, min(cap_cont, len(nomes)))
                # janelas de N dias: cada janela fixa um conjunto rotativo de alunos no serviço
                chunks = list(range(0, 5, dias_consec))
                for ci, cstart in enumerate(chunks):
                    desloc = (ci * cap_cont) % max(len(nomes), 1)
                    rot = nomes[desloc:] + nomes[:desloc]
                    presos = set(rot[:cap_cont])
                    for wd in range(cstart, min(cstart + dias_consec, 5)):
                        pinned[wd] = presos

            horas_aluno = {n: 0.0 for n in nomes}
            ocupado = {}  # (dia3,turno_key) -> set
            assigned = [[] for _ in slots]

            def _cabe(n, idx, hrs, dia3, tk, cap, respeitar_cont=True):
                if n in assigned[idx] or n in ocupado.get((dia3, tk), set()):
                    return False
                if horas_aluno[n] + hrs > cap:
                    return False
                # plantão exclusivo: n precisa já estar no serviço de origem nesta semana
                # (só para os turnos marcados; vazio = todos os turnos do serviço)
                _alvo_si = _excl_por_si.get(slot_svc[idx])
                if _alvo_si is not None:
                    _et = _excl_turnos_por_si.get(slot_svc[idx]) or set()
                    _cat = ("fds" if dia3 in ("Sáb", "Sab", "Dom") else "du") + "_" + tk
                    if (not _et) or (_cat in _et):
                        if not any(slot_svc[_j] == _alvo_si and n in assigned[_j] for _j in range(len(slots))):
                            return False
                if respeitar_cont and usa_continuidade and dia3 in dias3[:5]:  # só dias úteis
                    # a trava vale SÓ para os turnos do serviço de continuidade:
                    # ali só entram os alunos "presos" da janela; os demais serviços ficam livres
                    if slot_svc[idx] == cont_idx:
                        wd = dias3.index(dia3)
                        if n not in pinned.get(wd, set()):
                            return False
                return True

            def _por(n, idx, hrs, dia3, tk):
                assigned[idx].append(n)
                ocupado.setdefault((dia3, tk), set()).add(n)
                horas_aluno[n] += hrs
                g_tcount.setdefault(n, {})
                g_tcount[n][tk] = g_tcount[n].get(tk, 0) + 1

            def _tirar(n, idx, hrs, dia3, tk):
                assigned[idx].remove(n)
                ocupado.get((dia3, tk), set()).discard(n)
                horas_aluno[n] -= hrs
                if g_tcount.get(n, {}).get(tk):
                    g_tcount[n][tk] -= 1

            # ordena candidatos: menos turnos DESSE tipo no semestre primeiro, depois menos horas na semana
            def _ordem(tk):
                return lambda x: (g_tcount.get(x, {}).get(tk, 0), horas_aluno[x])

            # Preferência SUAVE por períodos consecutivos: quem já está de manhã no MESMO serviço/dia
            # tem leve preferência p/ a tarde; quem está de tarde, p/ a cinderela/noite. É só desempate
            # (entra DEPOIS do equilíbrio de turnos), então não atropela regras mais importantes.
            def _consec_pref(n, idx):
                tk = slots[idx][3]
                if tk == "tarde":
                    prev = "manha"
                elif tk == "cind":
                    prev = "tarde"
                else:
                    return 0
                dia = slots[idx][1]; si = slot_svc[idx]
                for jx in range(len(slots)):
                    if (slot_svc[jx] == si and slots[jx][1] == dia
                            and slots[jx][3] == prev and n in assigned[jx]):
                        return 0   # tem o período anterior no mesmo serviço/dia → preferido
                return 1

            def _ordem_ctx(idx, tk):
                return lambda x: (g_tcount.get(x, {}).get(tk, 0), _consec_pref(x, idx), horas_aluno[x])

            def _eh_fds_dia(d):
                return d in ("Sáb", "Sab", "Dom")

            # prioridade dos plantões de FDS: manhã/tarde do fim de semana primeiro (garante ≥1),
            # depois os dias úteis, e a CINDERELA de FDS por ÚLTIMO (só se sobrar/precisar de CH).
            def _prio_slot(ix):
                d = slots[ix][1]; tk = slots[ix][3]
                # serviços exclusivos por ÚLTIMO dentro do bucket: o serviço de origem
                # precisa já ter alunos antes de filtrar quem pode pegar o plantão exclusivo.
                excl = 1 if slot_svc[ix] in _excl_por_si else 0
                if _eh_fds_dia(d) and tk != "cind":
                    return (0, excl)
                if _eh_fds_dia(d) and tk == "cind":
                    return (2, excl)   # cinderela de FDS = última prioridade
                return (1, excl)       # dias úteis

            # Fase 1 — cobre o mínimo de cada turno (cobertura garantida), sem estourar.
            # Cobre manhã/tarde do FDS ANTES dos dias úteis; cinderela de FDS fica por último.
            ordem_fase1 = sorted(range(len(slots)), key=_prio_slot)
            for idx in ordem_fase1:
                (data, dia3, tn, tk, hor, hrs, qmin, qmax, locn) = slots[idx]
                for cap in (limite, limite_abs):
                    for n in sorted(nomes, key=_ordem_ctx(idx, tk)):
                        if len(assigned[idx]) >= qmin:
                            break
                        if _cabe(n, idx, hrs, dia3, tk, cap):
                            _por(n, idx, hrs, dia3, tk)
                    if len(assigned[idx]) >= qmin:
                        break

            # Fase 2 — completa quem está abaixo do alvo (sem passar de 40h), equilibrando por tipo de turno.
            # Usa a cinderela de FDS por ÚLTIMO (prioriza manhã/tarde para fechar a CH).
            ordem_fase2 = sorted(range(len(slots)), key=_prio_slot)
            progresso = True
            while progresso and any(horas_aluno[n] < alvo_min for n in nomes):
                progresso = False
                for idx in ordem_fase2:
                    (data, dia3, tn, tk, hor, hrs, qmin, qmax, locn) = slots[idx]
                    teto = qmax if qmax is not None else len(nomes)
                    while len(assigned[idx]) < teto:
                        cands = [n for n in nomes if horas_aluno[n] < alvo_min
                                 and _cabe(n, idx, hrs, dia3, tk, limite)]
                        if not cands:
                            break
                        n = min(cands, key=_ordem_ctx(idx, tk))
                        _por(n, idx, hrs, dia3, tk)
                        progresso = True

            # Fase 3 — reequilíbrio: transfere turnos de quem está ACIMA do alvo para quem está ABAIXO
            ciclos = 0
            while ciclos < 300:
                ciclos += 1
                pobres = sorted([n for n in nomes if horas_aluno[n] < alvo_min], key=lambda x: horas_aluno[x])
                if not pobres:
                    break
                movimentou = False
                for p in pobres:
                    # na fase de reequilíbrio a continuidade é RELAXADA (a CH alvo tem prioridade);
                    # cinderela de FDS continua sendo a última opção
                    for idx in ordem_fase2:
                        (data, dia3, tn, tk, hor, hrs, qmin, qmax, locn) = slots[idx]
                        if not _cabe(p, idx, hrs, dia3, tk, limite, respeitar_cont=False):
                            continue
                        teto = qmax if qmax is not None else len(nomes)
                        if len(assigned[idx]) < teto:
                            _por(p, idx, hrs, dia3, tk); movimentou = True; break
                        # slot cheio: troca por um colega "rico" que continua >= alvo após sair
                        ricos = [r for r in assigned[idx] if r != p
                                 and horas_aluno[r] - hrs >= alvo_min
                                 and horas_aluno[r] - hrs >= horas_aluno[p] + hrs]
                        if ricos:
                            r = max(ricos, key=lambda x: horas_aluno[x])
                            _tirar(r, idx, hrs, dia3, tk)
                            _por(p, idx, hrs, dia3, tk)
                            movimentou = True; break
                    if movimentou:
                        break
                if not movimentou:
                    break

            # ── Priorizar período (por serviço): (1) EQUILIBRA o período prioritário entre os
            # dias (evita amontoar muitos alunos num dia só); (2) o período SECUNDÁRIO só fica
            # com quem fez o prioritário no mesmo dia (os demais ficam com esse período livre =
            # Área Verde). priorizar_periodo: "manha" (secundário=tarde) ou "tarde" (sec.=manha).
            _prio_por_si = {}
            for _si2 in range(len(servs_bloco)):
                _p = servs_bloco[_si2].get("priorizar_periodo") or \
                     ("manha" if servs_bloco[_si2].get("priorizar_manha") else "")
                if _p in ("manha", "tarde"):
                    _prio_por_si[_si2] = _p

            # ── Equilíbrio por DIA (geral): evita amontoar alunos num dia (ex.: 6 na segunda)
            # e deixar outro no mínimo (ex.: 3 na sexta). Move do dia mais cheio para o mais
            # vazio dentro do MESMO serviço/turno, só nos dias úteis, mantendo o mínimo da origem
            # e o máximo do destino. CH não muda (mesmo turno/horas) — move-se primeiro p/ liberar
            # as horas e o _cabe avaliar o destino corretamente.
            for _si2 in range(len(servs_bloco)):
                if usa_continuidade and _si2 == cont_idx:
                    continue   # não mexe no serviço de continuidade (quebraria a sequência)
                for _tkb in ("manha", "tarde", "cind"):
                    if _prio_por_si.get(_si2) == _tkb:
                        continue   # período prioritário tem balanceamento próprio (abaixo)
                    _idxs = [ix for ix in range(len(slots))
                             if slot_svc[ix] == _si2 and slots[ix][3] == _tkb
                             and not _eh_fds_dia(slots[ix][1])]
                    if len(_idxs) < 2:
                        continue
                    for _ in range(120):
                        _idxs.sort(key=lambda ix: len(assigned[ix]))
                        _dst, _src = _idxs[0], _idxs[-1]
                        if len(assigned[_src]) - len(assigned[_dst]) <= 1:
                            break
                        if len(assigned[_src]) <= (slots[_src][6] or 0):   # não derruba o mínimo da origem
                            break
                        _qmx = slots[_dst][7]
                        if _qmx is not None and len(assigned[_dst]) >= _qmx:
                            break
                        _moved = False
                        for _n in list(assigned[_src]):
                            _tirar(_n, _src, slots[_src][5], slots[_src][1], _tkb)
                            if _cabe(_n, _dst, slots[_dst][5], slots[_dst][1], _tkb, limite, respeitar_cont=False):
                                _por(_n, _dst, slots[_dst][5], slots[_dst][1], _tkb)
                                _moved = True
                                break
                            _por(_n, _src, slots[_src][5], slots[_src][1], _tkb)   # desfaz
                        if not _moved:
                            break

            # (1) equilibra o período prioritário entre os dias (move dos dias mais cheios p/ os mais vazios)
            for _si2, _p in _prio_por_si.items():
                _idxs = [ix for ix in range(len(slots)) if slot_svc[ix] == _si2 and slots[ix][3] == _p]
                for _ in range(80):
                    _idxs.sort(key=lambda ix: len(assigned[ix]))
                    _dst, _src = _idxs[0], _idxs[-1]
                    if len(assigned[_src]) - len(assigned[_dst]) <= 1:
                        break
                    _qmx = slots[_dst][7]
                    if _qmx is not None and len(assigned[_dst]) >= _qmx:
                        break
                    _moved = False
                    for _n in list(assigned[_src]):
                        if _cabe(_n, _dst, slots[_dst][5], slots[_dst][1], _p, limite, respeitar_cont=False):
                            _tirar(_n, _src, slots[_src][5], slots[_src][1], _p)
                            _por(_n, _dst, slots[_dst][5], slots[_dst][1], _p)
                            _moved = True
                            break
                    if not _moved:
                        break

            # (2) o período secundário só fica com quem fez o prioritário no mesmo dia/serviço
            for _ix in range(len(slots)):
                _prio = _prio_por_si.get(slot_svc[_ix])
                if not _prio:
                    continue
                _sec = "tarde" if _prio == "manha" else "manha"
                if slots[_ix][3] != _sec:
                    continue
                _data_t = slots[_ix][0]
                _prim = set()
                for _jx in range(len(slots)):
                    if slot_svc[_jx] == slot_svc[_ix] and slots[_jx][3] == _prio and slots[_jx][0] == _data_t:
                        _prim |= set(assigned[_jx])
                for _n in [x for x in assigned[_ix] if x not in _prim]:
                    _tirar(_n, _ix, slots[_ix][5], slots[_ix][1], _sec)

            for idx, (data, dia3, tn, tk, hor, hrs, qmin, qmax, locn) in enumerate(slots):
                ch = assigned[idx]
                if not ch:
                    continue
                detalhada.append({
                    "semana": w, "data": data.strftime("%d/%m"), "dia": dia3,
                    "local": locn or b["nome"], "turno": tn, "horario": hor, "horas": hrs,
                    "sg": "+".join(sorted({sg_de[n] for n in ch}, key=lambda x: int(x) if x.isdigit() else 99)),
                    "alunos": ch,
                })

        # ── Complemento de CH: plantões em OUTRO serviço (ex.: Ambulatório → plantão no PA) ──
        # Para alunos de um bloco que não fecha a CH alvo, adiciona plantões em vagas LIVRES de
        # outro serviço, SEMPRE respeitando o máximo de alunos por turno/dia do serviço de destino.
        blocos_comp = [bb for bb in blocos if bb["loc"].get("complemento_ativo")]
        if blocos_comp:
            ent_sem = [e for e in detalhada if e.get("semana") == w]
            ha = {}                       # horas por aluno nesta semana
            occ = {}                      # (dia3, tk) -> set(alunos ocupados)
            for e in ent_sem:
                tkk = _turno_key(e.get("turno"))
                for al in e.get("alunos", []):
                    ha[al] = ha.get(al, 0.0) + _horas_entrada(e)
                    occ.setdefault((e.get("dia"), tkk), set()).add(al)

            def _rotulos(s):
                return {_norm(s.get("nome")), _norm(s.get("abrev"))} - {""}

            def _serv_por_nome(nome):
                a = _norm(nome)
                for bb in blocos:
                    for s in bb["servs"]:
                        rot = _rotulos(s)
                        if any(r and (r == a or r in a or a in r) for r in rot):
                            return s
                return None

            def _occ_serv(s, d3, tkk):  # nº de alunos já nesse serviço/dia/turno (respeitar máx)
                rot = _rotulos(s)
                c = 0
                for e in ent_sem:
                    if e.get("dia") != d3 or _turno_key(e.get("turno")) != tkk:
                        continue
                    ln = _norm(e.get("local"))
                    if any(r and (r in ln or ln in r) for r in rot):
                        c += len(e.get("alunos", []))
                return c

            for b2 in blocos_comp:
                rot_b2 = {_norm(b2["nome"])}
                for s in b2["servs"]:
                    rot_b2 |= _rotulos(s)
                alunos_b2 = set()
                for e in ent_sem:
                    if e.get("plantao"):
                        continue
                    ln = _norm(e.get("local"))
                    if any(r and (r in ln or ln in r) for r in rot_b2):
                        alunos_b2.update(e.get("alunos", []))
                if not alunos_b2:
                    continue
                # serviços-alvo: manual = o escolhido; auto (IA) = todos os de OUTROS blocos
                modo = b2["loc"].get("complemento_modo", "auto")
                alvos = []
                if modo == "manual" and b2["loc"].get("complemento_servico"):
                    sa = _serv_por_nome(b2["loc"]["complemento_servico"])
                    if sa is not None:
                        alvos = [sa]
                else:
                    for bb in blocos:
                        if bb is b2:
                            continue
                        alvos += bb["servs"]
                slots_alvo = []
                for sa in alvos:
                    for sl in _slots_servico(sa, datas):
                        slots_alvo.append((sa,) + tuple(sl))
                # Restringe aos períodos escolhidos para o complemento (DU/FDS × manhã/tarde/cinderela).
                _per_ok = b2["loc"].get("complemento_periodos")
                if _per_ok:
                    _per_ok = set(_per_ok)
                    def _cat_slot(sl):
                        _pre = "fds" if sl[2] in ("Sáb", "Sab", "Dom") else "du"
                        return f"{_pre}_{sl[4]}"
                    slots_alvo = [sl for sl in slots_alvo if _cat_slot(sl) in _per_ok]
                obrig = bool(b2["loc"].get("complemento_obrigatorio"))
                compensa = bool(b2["loc"].get("complemento_compensa"))

                def _min_blk_turno(local, tkk):
                    ln = _norm(local)
                    for s in b2["servs"]:
                        if any(r and (r in ln or ln in r) for r in _rotulos(s)):
                            return int(s.get(f"min_{tkk}", 0) or 0)
                    return 0

                comp_load = {}   # (dia3, tk) -> nº de plantões já colocados (p/ ESPALHAR entre dias/turnos)
                for al in sorted(alunos_b2, key=lambda x: ha.get(x, 0)):
                    # turnos já ocupados por dia (p/ ESPALHAR os plantões e evitar dia sobrecarregado)
                    day_load = {}
                    for (d3, _tk2), s_al in occ.items():
                        if al in s_al:
                            day_load[d3] = day_load.get(d3, 0) + 1
                    feitos = 0   # plantões de complemento dados a ESTE aluno nesta semana
                    # espalha: prioriza (1) o dia/turno MENOS usado pelos plantões e (2) o dia mais leve do aluno
                    ordenados = sorted(slots_alvo,
                                       key=lambda sl: (comp_load.get((sl[2], sl[4]), 0), day_load.get(sl[2], 0)))
                    for (sa, data, dia3, tn, tk, hor, hrs, qmin, qmax, locn) in ordenados:
                        # obrigatório: garante 1 plantão por aluno mesmo já tendo batido a CH;
                        # depois desse 1, volta a parar ao atingir a CH alvo.
                        if (not obrig or feitos >= 1) and ha.get(al, 0) >= alvo_min:
                            break
                        if al in occ.get((dia3, tk), set()):
                            continue
                        if day_load.get(dia3, 0) >= 2:   # no máx 2 turnos/dia (evita dia de 14h+)
                            continue
                        # no plantão obrigatório, permite chegar até o limite ABSOLUTO de CH
                        _cap_ch = limite_abs if (obrig and feitos < 1) else limite
                        if ha.get(al, 0) + hrs > _cap_ch:
                            continue
                        teto = qmax if qmax is not None else 99
                        if _occ_serv(sa, dia3, tk) >= teto:   # RESPEITA o máximo de alunos/turno
                            continue
                        occ.setdefault((dia3, tk), set()).add(al)
                        day_load[dia3] = day_load.get(dia3, 0) + 1
                        comp_load[(dia3, tk)] = comp_load.get((dia3, tk), 0) + 1
                        ha[al] = ha.get(al, 0) + hrs
                        feitos += 1
                        ent = {"semana": w, "data": data.strftime("%d/%m"), "dia": dia3,
                               "local": locn, "turno": tn, "horario": hor, "horas": hrs,
                               "sg": sg_de_aluno.get(al, ""), "alunos": [al], "plantao": True}
                        detalhada.append(ent)
                        ent_sem.append(ent)

                    # Compensação: quem fez o plantão obrigatório tem 1 período de DIA ÚTIL retirado
                    # (preferindo a tarde), desde que o serviço siga com o mínimo e o aluno com a CH mínima.
                    if obrig and compensa and feitos >= 1:
                        cands = []
                        for e in ent_sem:
                            if e.get("plantao") or e.get("dia") in ("Sáb", "Sab", "Dom"):
                                continue
                            if al not in e.get("alunos", []):
                                continue
                            ln = _norm(e.get("local"))
                            if not any(r and (r in ln or ln in r) for r in rot_b2):
                                continue
                            tkk = _turno_key(e.get("turno"))
                            mn = _min_blk_turno(e.get("local"), tkk)
                            if len(e.get("alunos", [])) - 1 < mn:        # mantém o mínimo do serviço
                                continue
                            hh = _horas_entrada(e)
                            if ha.get(al, 0) - hh < alvo_min:            # não derruba abaixo da CH mínima
                                continue
                            prio = 0 if tkk == "tarde" else (1 if tkk == "manha" else 2)
                            cands.append((prio, -(len(e["alunos"]) - mn), hh, e, tkk))
                        if cands:
                            cands.sort(key=lambda x: (x[0], x[1]))
                            _, _, hh, e, tkk = cands[0]
                            e["alunos"].remove(al)
                            ha[al] = ha.get(al, 0) - hh
                            occ.get((e.get("dia"), tkk), set()).discard(al)
                            if not e["alunos"]:   # não deixa linha fantasma sem alunos
                                for _lst in (detalhada, ent_sem):
                                    try:
                                        _lst.remove(e)
                                    except ValueError:
                                        pass
    detalhada.sort(key=lambda e: (e["semana"], e["data"], e["local"], e["turno"]))
    return detalhada

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
    if val.get("nao_passou"):
        st.warning(f"🔁 **{len(val['nao_passou'])} aluno(s) NÃO passaram por algum serviço (rodízio incompleto):**")
        for np_ in val["nao_passou"][:25]:
            st.markdown(f"- **{np_['aluno']}**: faltou passar em {', '.join(np_['faltam'])}")
        if len(val["nao_passou"]) > 25:
            st.caption(f"...e mais {len(val['nao_passou']) - 25}")
        st.caption("💡 Geralmente é a tabela de rodízio (Bloco 6): ajuste para que todos os SGs passem por todos os locais.")
    elif val.get("pct_servico"):
        st.success("✅ Todos os alunos foram alocados em **todos os serviços**.")

    # ── % de tempo de cada aluno em cada serviço (visual) ────────────────────
    if val.get("pct_servico") and val.get("locais_lista"):
        locs = val["locais_lista"]
        pcts = val["pct_servico"]
        # destaca quem passa pouco tempo (>0% e <15%) em algum serviço
        pouco = []
        for al, d in pcts.items():
            for loc in locs:
                if 0 < d.get(loc, 0) < 15:
                    pouco.append(f"{al} — {loc}: {d[loc]}%")
        with st.expander("📊 % de tempo de cada aluno por serviço", expanded=bool(pouco)):
            rows = []
            for al in sorted(pcts):
                row = {"Aluno": al}
                for loc in locs:
                    p = pcts[al].get(loc, 0)
                    row[loc] = f"{p}%" if p else "—"
                rows.append(row)
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
            if pouco:
                st.warning("⚠️ **Pouco tempo (<15%) em algum serviço:**\n- " + "\n- ".join(pouco[:20]))
            st.caption("'—' = não passou nesse serviço. % = fração da carga horária total do aluno naquele serviço.")
    if val.get("sobrecarga_bloco"):
        alvo = val.get("alvo_min", 34)
        st.error(f"🚨 **Bloco lotado em alguma semana (alunos demais p/ a capacidade → CH não fecha):**")
        for sb in val["sobrecarga_bloco"][:20]:
            st.markdown(f"- **{sb['bloco']}** — semana {sb['semana']}: **{sb['alunos']} alunos** ao mesmo tempo; "
                        f"o bloco só comporta ~**{sb['ch_max']}h/aluno** (alvo {alvo}h)")
        st.caption("💡 **Esta é a causa da CH baixa.** No rodízio (Bloco 6), evite colocar tantos subgrupos no "
                   "mesmo bloco na mesma semana — distribua-os por outros blocos. Ou aumente a capacidade do "
                   "bloco (mais turnos/serviços ou maior máx/dia).")
    if val.get("plantoes_complemento"):
        st.info(f"🩺 **{val['plantoes_complemento']} plantão(ões) de complemento** adicionados em outro serviço "
                f"para completar a carga horária — respeitando o **máximo de alunos por turno/dia** do serviço de destino.")
    if val.get("subcarga"):
        alvo = val.get("alvo_min", 34)
        st.warning(f"⏬ **{len(val['subcarga'])} aluno(s)/semana ABAIXO da CH mínima alvo ({alvo}h):**")
        for sgc in val["subcarga"][:20]:
            st.markdown(f"- **{sgc['aluno']}** — semana {sgc['semana']}: só **{sgc['horas']}h** (alvo {alvo}h)")
        if len(val["subcarga"]) > 20:
            st.caption(f"...e mais {len(val['subcarga']) - 20}")
        st.caption("💡 Para subir a CH: ative o **complemento com plantões em outro serviço** (Bloco 5), "
                   "ative mais turnos no bloco (ex: cinderela), aumente o máx/dia dos turnos, "
                   "ou reduza a CH mínima alvo no Bloco 3. Se já usa complemento e ainda falta, "
                   "o serviço de destino atingiu o **máximo de alunos** — aumente o máx/dia dele.")

# ── Pedidos extraordinários: avaliação e trocas automáticas ──────────────────
import unicodedata as _ud

def _sem_acento(s):
    return "".join(c for c in _ud.normalize("NFKD", str(s or "")) if not _ud.combining(c)).strip().lower()

def _data_ddmm(s):
    """Normaliza 'DD/MM/AAAA' ou 'DD/MM' -> 'DD/MM'."""
    p = str(s or "").strip().split("/")
    if len(p) >= 2:
        try:
            return f"{int(p[0]):02d}/{int(p[1]):02d}"
        except ValueError:
            pass
    return str(s or "").strip()

def _extrair_datas(texto):
    """Extrai todas as datas DD/MM ou DD/MM/AAAA de um texto (ignora anotações como '(Sex)')."""
    return _re_val.findall(r"\d{1,2}/\d{1,2}(?:/\d{2,4})?", str(texto or ""))

def _roster_nomes(dados):
    nomes = set()
    for e in (dados.get("escala_detalhada") or []):
        for a in _alunos_entrada(e):
            nomes.add(a)
    return sorted(nomes)

def _casar_aluno(nome_pedido, roster):
    """Acha na escala o nome que melhor casa com o do pedido (sem acento, por tokens)."""
    toks_alvo = _sem_acento(nome_pedido).split()
    alvo = set(toks_alvo)
    if not alvo:
        return None
    primeiro = toks_alvo[0]
    need = min(2, len(toks_alvo))   # nomes de 1 palavra casam por 1 token
    melhor, score = None, -1
    for nome in roster:
        toks = set(_sem_acento(nome).split())
        inter = len(alvo & toks)
        if inter < need:
            continue
        s = inter + (0.5 if primeiro in toks else 0)
        if s > score:
            score, melhor = s, nome
    return melhor

def avaliar_pedidos_extraordinarios(dados, config, pedidos):
    """Para cada (aluno, data) diz se o aluno está escalado e onde.
    pedidos: [{'aluno': str, 'datas': [str]}]. Retorna lista de registros."""
    det = dados.get("escala_detalhada") or []
    roster = _roster_nomes(dados)
    idx = _dd(list)
    for i, e in enumerate(det):
        dd = _data_ddmm(e.get("data"))
        for a in _alunos_entrada(e):
            idx[(a, dd)].append(i)
    datas_escala = {_data_ddmm(e.get("data")) for e in det}
    linhas = []
    for p in pedidos:
        nome_real = _casar_aluno(p.get("aluno", ""), roster)
        for d in p.get("datas", []):
            dd = _data_ddmm(d)
            reg = {"aluno_pedido": p.get("aluno", ""), "aluno_escala": nome_real or "(não encontrado)",
                   "data": dd, "no_periodo": dd in datas_escala, "escalado": False, "entradas": []}
            if nome_real:
                for i in idx.get((nome_real, dd), []):
                    e = det[i]
                    reg["escalado"] = True
                    reg["entradas"].append({"idx": i, "local": e.get("local", ""), "turno": e.get("turno", ""),
                                            "horario": e.get("horario", ""), "dia": e.get("dia", ""),
                                            "semana": e.get("semana", ""), "horas": _horas_entrada(e)})
            linhas.append(reg)
    return linhas

def _faixa_servico(config, local, turno_key, dia):
    """(min, max) de alunos para um local/turno/dia, do Bloco 5."""
    eh_fds = _dia_curto(dia) in ("Sab", "Dom")
    chave = "turnos_fds" if eh_fds else "turnos"
    for s in _servicos_config(config):
        if turno_key in s.get(chave, {}) and _nome_bate(s["nomes"], _norm(local)):
            return s[chave][turno_key]
    return (0, None)

def _ch_semana(det, al, sem):
    return sum(_horas_entrada(e) for e in det if e.get("semana") == sem and al in _alunos_entrada(e))

def _conta_turno(det, local, turno, sem, dia):
    tk = _turno_key(turno)
    return sum(len(_alunos_entrada(e)) for e in det
               if e.get("local") == local and _turno_key(e.get("turno")) == tk
               and e.get("semana") == sem and _dia_curto(e.get("dia")) == _dia_curto(dia))

def _achar_substituto(det, config, al_fora, local, turno, sem, dia, horas, limite_abs):
    """Colega da coorte do serviço, livre nesse turno/dia, que cabe sem furar máx nem CH."""
    tk = _turno_key(turno)
    _, mx = _faixa_servico(config, local, tk, dia)
    ocupados = set()
    for e in det:
        if (e.get("semana") == sem and _dia_curto(e.get("dia")) == _dia_curto(dia)
                and _turno_key(e.get("turno")) == tk):
            ocupados.update(_alunos_entrada(e))
    coorte = set()
    for e in det:
        if e.get("local") == local:
            coorte.update(_alunos_entrada(e))
    cand = [c for c in coorte if c != al_fora and c not in ocupados
            and _ch_semana(det, c, sem) + horas <= limite_abs]
    if mx is not None and _conta_turno(det, local, turno, sem, dia) >= mx:
        return None
    cand.sort(key=lambda c: _ch_semana(det, c, sem))
    return cand[0] if cand else None

def _achar_reposicao(det, config, al, local, turno_perdido, sem, limite_abs, datas_proibidas=None):
    """Outra entrada no MESMO local, em DIA diferente dos pedidos, onde o aluno cabe
    (livre no turno, vaga < máx, CH ok). datas_proibidas: set de 'DD/MM' a evitar."""
    tk_perd = _turno_key(turno_perdido)
    datas_proibidas = datas_proibidas or set()
    ocup = set()
    for e in det:
        if al in _alunos_entrada(e):
            ocup.add((e.get("semana"), _dia_curto(e.get("dia")), _turno_key(e.get("turno"))))
    melhor = None
    for i, e in enumerate(det):
        if e.get("local") != local:
            continue
        if al in _alunos_entrada(e):
            continue
        if _data_ddmm(e.get("data")) in datas_proibidas:   # nunca repor num dia pedido
            continue
        sem_e, dia_e = e.get("semana"), _dia_curto(e.get("dia"))
        tk_e = _turno_key(e.get("turno"))
        if (sem_e, dia_e, tk_e) in ocup:
            continue
        _, mx_e = _faixa_servico(config, local, tk_e, e.get("dia"))
        n = len(_alunos_entrada(e))
        if mx_e is not None and n >= mx_e:
            continue
        if _ch_semana(det, al, sem_e) + _horas_entrada(e) > limite_abs:
            continue
        score = (0 if tk_e == tk_perd else 1, n)   # mesmo turno primeiro, depois menos cheio
        if melhor is None or score < melhor[0]:
            melhor = (score, i)
    return melhor[1] if melhor else None

def aplicar_trocas_pedidos(dados, config, avaliacao):
    """Dispensa o aluno dos dias pedidos e tenta repor o serviço em outro dia,
    respeitando mín/máx por local/turno (Bloco 5) e a CH absoluta. Retorna (dados_novo, relatorio)."""
    import copy
    dados = copy.deepcopy(dados)
    det = dados.get("escala_detalhada") or []
    limite_abs = int(config.get("regras_especiais", {}).get("limite_abs", 43))
    # datas que cada aluno pediu para NÃO ser escalado (nunca repor nesses dias)
    proibidas = _dd(set)
    for reg in avaliacao:
        proibidas[reg["aluno_escala"]].add(_data_ddmm(reg["data"]))
    relatorio = []
    for reg in avaliacao:
        if not reg.get("escalado"):
            continue
        al = reg["aluno_escala"]
        for ent in reg["entradas"]:
            i = ent["idx"]
            e = det[i]
            local, turno = e.get("local", ""), e.get("turno", "")
            sem, dia = e.get("semana", ""), e.get("dia", "")
            horas = _horas_entrada(e)
            tk = _turno_key(turno)
            mn, _ = _faixa_servico(config, local, tk, dia)
            antes = _conta_turno(det, local, turno, sem, dia)
            # 1) dispensa
            e["alunos"] = [x for x in _alunos_entrada(e) if x != al]
            e.pop("nome", None)
            depois = antes - 1
            item = {"aluno": al, "data": reg["data"], "dia": _dia_curto(dia), "servico": f"{local} / {turno}",
                    "substituto": "—", "reposicao": "—",
                    "cobertura": f"{antes}→{depois} (mín {mn})", "status": "ok"}
            # 2) cobertura mínima
            if depois < mn:
                sub = _achar_substituto(det, config, al, local, turno, sem, dia, horas, limite_abs)
                if sub:
                    det[i]["alunos"] = _alunos_entrada(det[i]) + [sub]
                    item["substituto"] = sub
                    item["cobertura"] = f"{antes}→{depois}→{depois+1} (mín {mn})"
                else:
                    item["status"] = "atenção: cobertura abaixo do mínimo (sem substituto livre)"
            # 3) reposição no mesmo serviço
            j = _achar_reposicao(det, config, al, local, turno, sem, limite_abs, proibidas.get(al))
            if j is not None:
                det[j]["alunos"] = _alunos_entrada(det[j]) + [al]
                item["reposicao"] = f'{_data_ddmm(det[j].get("data"))} ({_dia_curto(det[j].get("dia"))}) {det[j].get("turno")}'
            else:
                item["reposicao"] = "pendente (sem vaga no mesmo serviço)"
                if item["status"] == "ok":
                    item["status"] = "atenção: reposição pendente"
            relatorio.append(item)
    dados["escala_detalhada"] = det
    dados["resumo_horas"] = recalcular_resumo_horas(dados, config)
    return dados, relatorio

def _parse_pedidos_editor(df):
    """DataFrame com colunas Aluno / Datas -> [{'aluno','datas':[...]}]."""
    pedidos = []
    for _, row in df.iterrows():
        aluno = str(row.get("Aluno", "") or "").strip()
        datas_raw = str(row.get("Datas", "") or "").strip()
        if not aluno or not datas_raw:
            continue
        datas = _extrair_datas(datas_raw)
        if datas:
            pedidos.append({"aluno": aluno, "datas": datas})
    return pedidos

def _pedidos_de_arquivo(file):
    """Lê CSV/XLSX de pedidos. Aceita coluna de aluno + 1+ colunas de data
    (uma data por linha OU várias datas na mesma célula). Agrupa por aluno."""
    try:
        if file.name.lower().endswith(".csv"):
            df = pd.read_csv(file, dtype=str, keep_default_na=False, sep=None, engine="python")
        else:
            df = pd.read_excel(file, dtype=str)
    except Exception:
        return None
    df.columns = [str(c).strip() for c in df.columns]
    col_aluno = next((c for c in df.columns
                      if "aluno" in _sem_acento(c) or "nome" in _sem_acento(c)), None)
    cols_data = [c for c in df.columns if "data" in _sem_acento(c)]
    if not col_aluno or not cols_data:
        return None
    agrup = {}
    for _, row in df.iterrows():
        al = str(row.get(col_aluno, "") or "").strip()
        if not al:
            continue
        datas = []
        for c in cols_data:
            v = str(row.get(c, "") or "").strip()
            if v and v.lower() != "nan":
                datas += _extrair_datas(v)
        agrup.setdefault(al, [])
        for d in datas:
            if d not in agrup[al]:
                agrup[al].append(d)
    return [{"aluno": k, "datas": v} for k, v in agrup.items() if v]

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

    # Exportar para o Lovable (2 formatos)
    st.markdown("**📤 Exportar para o Lovable**")
    cfg_lv = dict(st.session_state.get("config_atual", {}))
    cfg_lv.update({"especialidade": esp, "grupo": grupo, "turma": turma})
    _slug_lv = _re_val.sub(r"[^A-Za-z0-9]+", "_", f"{esp}_{turma}").strip("_") or "rodizio"
    colL1, colL2 = st.columns(2)
    try:
        from exportar_lovable import gerar_template_lovable, gerar_correcao_lovable
        with colL1:
            st.download_button(
                "🗂️ Arquivo 1 — Estrutura (blocos, semana padrão, rodízio, subgrupos)",
                data=gerar_template_lovable(dados, cfg_lv),
                file_name=f"template_{_slug_lv}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="Como funciona: blocos, semana padrão, rodízio e divisão dos subgrupos. "
                     "Abas: Semana Padrão · Distribuição · _Serviços · _Blocos · _Subgrupos · LEIA-ME.")
        with colL2:
            st.download_button(
                "📥 Arquivo 2 — Importar no Lovable (escala dia a dia)",
                data=gerar_correcao_lovable(dados, cfg_lv),
                file_name=f"importar_lovable_{_slug_lv}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="No formato exato de importação do Lovable: 1 linha por aluno × dia "
                     "(manhã/tarde/noite, com Área Verde nos períodos livres de dia útil).")
    except Exception as e:
        st.warning(f"Erro ao gerar os arquivos do Lovable: {e}")

    # ── Prompt pronto para publicar via Claude Code (copie e cole) ──
    _SIGLAS = {"GINECOLOGIA": "GO", "CLÍNICA MÉDICA": "CM", "CLINICA MEDICA": "CM",
               "CIRÚRGICA": "CIRURGIA", "CIRURGICA": "CIRURGIA", "CIRURGIA": "CIRURGIA",
               "PEDIATRIA": "PED", "FAMÍLIA": "MFC", "FAMILIA": "MFC", "MENTAL": "SM"}
    _sig = next((v for k, v in _SIGLAS.items() if k in (esp or "").upper()), "XX")
    _cod_sug = f"R1-{_sig}-{turma}" if turma else f"R1-{_sig}-T?"
    with st.expander("🤖 Publicar no Lovable via Claude Code (prompt pronto)", expanded=True):
        cpa, cpb = st.columns(2)
        with cpa:
            _cod = st.text_input("Código do rodízio", value=_cod_sug, key="_pub_codigo",
                help="Ex.: R1-GO-T6 (Rodízio 1 · GO · Turma 6). Ajuste o número do rodízio se preciso.")
        with cpb:
            _nsg_def = st.session_state.get("config_atual", {}).get("num_sg", 6)
            _nsg_opts = [4, 6, 8]
            _nsg = st.selectbox("Subgrupos (opcao_total_sg)", _nsg_opts,
                index=_nsg_opts.index(_nsg_def) if _nsg_def in _nsg_opts else 1, key="_pub_nsg")
        _prompt_pub = (
            f"Publique a escala {_cod} no Lovable usando publicar_lovable.py.\n"
            f"- Código do rodízio: {_cod}\n"
            f"- Subgrupos (opcao_total_sg): {_nsg}\n"
            f"- Especialidade: {esp} | Turma: {turma}\n"
            f"- Anexei os 3 arquivos do gerador: template_{_slug_lv}.xlsx, "
            f"importar_lovable_{_slug_lv}.xlsx e o definicoes_*.json.\n"
            f"Faça dry-run primeiro pra eu conferir, depois publique e valide "
            f"(CH, plano de blocos, rotação diagonal e área verde)."
        )
        st.caption("1) Baixe os 3 arquivos acima.  2) Copie este prompt (botão 📋 no canto).  "
                   "3) Abra uma conversa NOVA no Claude Code e cole o prompt + anexe os 3 arquivos.")
        st.code(_prompt_pub, language="text")

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

⚠️ BLOCO ≠ SERVIÇO (muito importante):
Um "bloco de rodízio" é uma ETAPA do rodízio pela qual cada subgrupo passa, e pode conter 1, 2 ou 3
SERVIÇOS que acontecem JUNTOS naquela etapa. Ex.: "Enfermaria Cirúrgica + Centro Cirúrgico" são 2
serviços do MESMO bloco; 3 estágios de Anestesiologia em hospitais diferentes que ocorrem no mesmo
período também podem ser 1 bloco com 3 serviços. NÃO crie um bloco separado para cada serviço.
- Agrupe no MESMO bloco os serviços que os alunos cursam ao mesmo tempo / no mesmo período do rodízio
  (em geral mesmo hospital ou mesma etapa). Use "nome_bloco" para o nome da etapa, os campos do
  serviço principal no próprio item, e "servicos_extras" para o 2º/3º serviço do mesmo bloco.
- "num_locais" = número de BLOCOS (etapas do rodízio), NÃO de serviços. Cada item de "locais" é 1 bloco.
- Se um bloco tem só 1 serviço, deixe "servicos_extras": [].
- REGRA OBRIGATÓRIA: se o nome de uma etapa juntar 2+ serviços com "e", "+", "/", "&" ou "com"
  (ex.: "Enfermaria Cirúrgica e Centro Cirúrgico"), crie CADA serviço SEPARADAMENTE no mesmo bloco —
  o 1º em "nome", os demais em "servicos_extras" — e ponha o nome composto em "nome_bloco". NUNCA
  deixe um nome composto inteiro dentro de um único campo "nome".
  Exemplo: bloco "Enfermaria Cirúrgica e Centro Cirúrgico" → nome_bloco="Enfermaria Cirúrgica e
  Centro Cirúrgico", nome="Enfermaria Cirúrgica", servicos_extras=[{{"nome":"Centro Cirúrgico", ...}}].
- AGRUPE por especialidade: várias etapas da MESMA especialidade em hospitais/locais diferentes que
  formam um único momento do rodízio devem virar UM bloco com vários serviços (o aluno roda
  internamente entre eles, e a duração do bloco = soma das semanas desses serviços). Exemplos reais
  de Cirurgia: bloco "Anestesiologia" = Anest. Santa Casa Leme + Anest. SCA + Anest. Mandic (3
  serviços); bloco "Ambulatório" = Amb. de Cirurgia + Amb. de Oftalmologia (2 serviços).
  ATENÇÃO: nem toda etapa de nome parecido se junta — Enfermarias em hospitais diferentes podem ser
  blocos separados se forem momentos distintos do rodízio. Na dúvida, agrupe só quando claramente é o
  mesmo bloco; o usuário ajusta o resto.

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
      "nome_bloco": "ex: Enfermaria Cirúrgica e Centro Cirúrgico",
      "nome": "", "abrev": "", "obs": "",
      "manha": "07-13h", "min_manha": 0, "max_manha": 6,
      "tarde": "12-18h", "min_tarde": 3, "max_tarde": 4,
      "bloqueios_tarde": [{{"dia": "Qui", "tipo": "Sem tarde", "horario": ""}}, {{"dia": "Ter", "tipo": "Horário reduzido", "horario": "12-16h"}}],
      "cinderela": "", "min_cind": 0, "max_cind": 0, "dias_cind": [],
      "fds": false, "fds_manha": "", "fds_tarde": "", "fds_cind": "",
      "fds_min_manha": 0, "fds_max_manha": 2, "fds_min_tarde": 0, "fds_max_tarde": 2,
      "fds_min_cind": 0, "fds_max_cind": 0,
      "fds_quem": "", "fds_comp": "",
      "servicos_extras": [
        {{"nome": "2º (ou 3º) serviço DESTE MESMO bloco — deixe esta lista vazia se o bloco tem só 1 serviço",
          "abrev": "", "obs": "",
          "manha": "07-13h", "min_manha": 0, "max_manha": 6,
          "tarde": "", "min_tarde": 0, "max_tarde": 0,
          "cinderela": "", "min_cind": 0, "max_cind": 0, "dias_cind": [],
          "fds": false, "fds_manha": "", "fds_tarde": "", "fds_cind": ""}}
      ],
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
                        # Limpa o cache das caixas de alunos (evita mostrar valores antigos/vazios)
                        for _k in [k for k in list(st.session_state.keys()) if str(k).startswith("sg_imp_")]:
                            del st.session_state[_k]
                        # Se tiver base de alunos, usar os nomes OFICIAIS do grupo da escala
                        # (a escala importada pode conter os alunos do grupo errado).
                        if arquivo_alunos_imp:
                            st.session_state.prefill["_arquivo_alunos"] = True
                            try:
                                _b = arquivo_alunos_imp.read()
                                _xls_b = pd.ExcelFile(io.BytesIO(_b), engine="openpyxl")
                                _ns = int(dados.get("num_sg", 6) or 6)
                                _gpname = str(dados.get("grupo", "")).strip()
                                _gp = _gpname.upper().replace("GRUPO", "").strip()
                                _novos, _ra, _shname = None, {}, None
                                _long_b = _abas_long_alunos(_xls_b)
                                if _long_b:
                                    # Formato tabela longa: escolhe grupo + aba/opção com nº de SG = _ns
                                    _grps = _grupos_long(_b, _long_b[0])
                                    _gmatch = next((g for g in _grps
                                                    if g.upper().replace("GRUPO", "").strip() == _gp), _gpname)
                                    _best = None
                                    for s in _long_b:
                                        _c = int(_df_long_grupo(_b, s, _gmatch)["Sub Grupo"].dropna().nunique())
                                        if _best is None:
                                            _best = s
                                        if _c == _ns:
                                            _best = s
                                            break
                                    if _best:
                                        _dff = _df_long_grupo(_b, _best, _gmatch)
                                        _novos, _ra = _ler_subgrupos(_dff)
                                        _shname = f"{_gmatch} ({_best})"
                                else:
                                    _sheets_g = [s for s in _xls_b.sheet_names if "GRUPO" in s.upper()]
                                    _sh = next((s for s in _sheets_g
                                                if s.upper().replace("GRUPO", "").strip() == _gp), None)
                                    if _sh:
                                        _dfg = pd.read_excel(io.BytesIO(_b), engine="openpyxl", sheet_name=_sh)
                                        if "OPÇÃO" in _dfg.columns:
                                            _ops = list(_dfg["OPÇÃO"].dropna().unique())
                                            _op = next((o for o in _ops if f"{_ns} SG" in str(o)), _ops[0] if _ops else None)
                                            if _op is not None:
                                                _dff = _dfg[_dfg["OPÇÃO"] == _op]
                                                _novos, _ra = _ler_subgrupos(_dff)
                                                _shname = _sh
                                if _novos:
                                    st.session_state.prefill["alunos_por_sg"] = _novos
                                    st.session_state.prefill["num_sg"] = len(_novos)
                                    st.session_state["ra_por_aluno"] = _ra
                                    st.session_state.prefill["_alunos_base_grupo"] = _shname
                            except Exception:
                                pass   # mantém os alunos extraídos da escala se a base falhar
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

# ── Catálogo de especialidades/serviços CADASTRADOS no Lovable (alimenta os seletores) ──
# Evita digitar serviço/local errado: o nome vem de uma lista, não de texto livre.
def _carregar_catalogo_lovable():
    import json as _json, os as _os
    _p = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "servicos_lovable.json")
    try:
        with open(_p, encoding="utf-8") as _f:
            return _json.load(_f)
    except Exception:
        return {"especialidades": [], "servicos_compartilhados": [], "servicos_por_especialidade": {}}

_CAT_LOVABLE = _carregar_catalogo_lovable()

def _servicos_da_especialidade(esp):
    """Serviços cadastrados da especialidade + os compartilhados (área verde, aulas)."""
    por = _CAT_LOVABLE.get("servicos_por_especialidade", {})
    alvo = _sem_acento(esp)
    nomes = []
    for k, v in por.items():
        if _sem_acento(k) == alvo:
            nomes = list(v)
            break
    return nomes + list(_CAT_LOVABLE.get("servicos_compartilhados", []))

st.header("📋 Briefing da Escala")
st.caption("Preencha com atenção — quanto mais detalhado, mais precisa a escala gerada.")

# BLOCO 1
with st.expander("📌 Bloco 1 — Identificação", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        _esp_cad = _CAT_LOVABLE.get("especialidades", [])
        _esp_cur = pf.get("especialidade", "")
        _ESP_OUTRA = "➕ Outra (digitar)…"
        _esp_opts = list(_esp_cad)
        if _esp_cur and _esp_cur not in _esp_opts:
            _esp_opts = [_esp_cur] + _esp_opts
        if _esp_cad:
            _esp_opts = _esp_opts + [_ESP_OUTRA]
            _esp_sel = st.selectbox("Especialidade *", _esp_opts,
                index=(_esp_opts.index(_esp_cur) if _esp_cur in _esp_opts else 0),
                help="Escolha da lista cadastrada no Lovable. Se faltar, use 'Outra (digitar)'.")
            if _esp_sel == _ESP_OUTRA:
                especialidade = st.text_input("Especialidade (digitar) *", value=_esp_cur, placeholder="ex: Clínica Médica")
            else:
                especialidade = _esp_sel
        else:
            # catálogo indisponível -> mantém digitação livre (comportamento antigo)
            especialidade = st.text_input("Especialidade *", value=_esp_cur, placeholder="ex: Clínica Médica")
        # opções de serviço da especialidade escolhida (usadas no seletor de cada serviço)
        st.session_state["_serv_opts_cad"] = _servicos_da_especialidade(especialidade)
        anos = ["3º Ano","4º Ano","5º Ano","6º Ano"]
        ano_idx = anos.index(pf.get("ano_curso","4º Ano")) if pf.get("ano_curso") in anos else 1
        ano_curso = st.selectbox("Ano do curso *", anos, index=ano_idx)
        turma = st.text_input("Turma *", value=pf.get("turma",""), placeholder="ex: T6")
    with col2:
        grupo = st.text_input("Grupo *", value=pf.get("grupo",""), placeholder="ex: Grupo A")
        try: data_def = datetime.date.fromisoformat(pf.get("data_inicio",""))
        except: data_def = datetime.date.today()
        data_inicio = st.date_input("Data de início (segunda-feira) *", value=data_def)
        num_semanas = st.number_input("Número de semanas *", 1, 20, _clamp(pf.get("num_semanas",8), 1, 20, 8))

# BLOCO 2
with st.expander("👥 Bloco 2 — Alunos e Subgrupos", expanded=True):
    arquivo_alunos = st.file_uploader("Upload Excel de alunos (opcional)", type=["xlsx"])
    num_sg = st.number_input("Número de subgrupos", 2, 8, _clamp(pf.get("num_sg",6), 2, 8, 6))
    alunos_por_sg = {}

    # Se a BASE de alunos foi enviada: mostra SEMPRE o seletor de Grupo (A/B/C...) + opção de
    # subgrupos e carrega os alunos ao vivo. Vale para "Criar nova" e para "Importar"
    # (substitui os alunos que vieram da escala importada).
    if arquivo_alunos:
        try:
            _bytes_alunos2 = arquivo_alunos.read()
            xls2 = pd.ExcelFile(io.BytesIO(_bytes_alunos2), engine="openpyxl")
            _long2 = _abas_long_alunos(xls2)
            grupos_disp = [s for s in xls2.sheet_names if "GRUPO" in s.upper()]
            if _long2:
                # ---- Formato tabela longa: 1 aba por opção de SG + coluna "Grupo" ----
                _grupos = _grupos_long(_bytes_alunos2, _long2[0])
                _gp = (grupo or pf.get("grupo", "") or "").strip().upper().replace("GRUPO", "").strip()
                _gidx = next((i for i, g in enumerate(_grupos)
                              if _gp and g.upper().replace("GRUPO", "").strip() == _gp), 0)
                col_g1, col_g2 = st.columns(2)
                with col_g1:
                    grupo_sel = st.selectbox("👥 Grupo (escolha o grupo certo)", _grupos,
                                             index=_gidx if _grupos else 0)
                # rótulo de cada opção com a contagem REAL de SG para este grupo
                _labels, _sheets = [], []
                for s in _long2:
                    _c = int(_df_long_grupo(_bytes_alunos2, s, grupo_sel)["Sub Grupo"].dropna().nunique())
                    _labels.append(f"{s} — {_c} SG")
                    _sheets.append((s, _c))
                _oidx = next((i for i, (_, c) in enumerate(_sheets) if c == int(num_sg)), 0)
                with col_g2:
                    _osel = st.selectbox("Opção de subgrupos", _labels,
                                         index=_oidx if _labels else 0)
                _sheet_sel = _sheets[_labels.index(_osel)][0]
                df_f = _df_long_grupo(_bytes_alunos2, _sheet_sel, grupo_sel)
                alunos_por_sg, ra_map = _ler_subgrupos(df_f)
                st.session_state["ra_por_aluno"] = ra_map
                st.success(f"✅ {grupo_sel}: {len(df_f)} alunos em {len(alunos_por_sg)} SGs (com RA). "
                           f"Primeiros: {', '.join(df_f['Nome Completo'].head(3).astype(str))}…")
            elif grupos_disp:
                # default já no grupo informado/importado (ex.: "Grupo C" -> aba "GRUPO C")
                _gp = (grupo or pf.get("grupo", "") or "").upper().replace("GRUPO", "").strip()
                _idx = next((i for i, s in enumerate(grupos_disp)
                             if _gp and s.upper().replace("GRUPO", "").strip() == _gp), 0)
                col_g1, col_g2 = st.columns(2)
                with col_g1:
                    grupo_sel = st.selectbox("👥 Grupo (escolha o grupo certo)", grupos_disp, index=_idx)
                df_g = pd.read_excel(io.BytesIO(_bytes_alunos2), engine="openpyxl", sheet_name=grupo_sel)
                if "OPÇÃO" in df_g.columns:
                    opcoes = list(df_g["OPÇÃO"].dropna().unique())
                    # default na opção que bate com o nº de subgrupos escolhido
                    _oidx = next((i for i, o in enumerate(opcoes) if f"{int(num_sg)} SG" in str(o)), 0)
                    with col_g2:
                        opcao_sel = st.selectbox("Opção de subgrupos", opcoes, index=_oidx)
                    df_f = df_g[df_g["OPÇÃO"] == opcao_sel]
                    alunos_por_sg, ra_map = _ler_subgrupos(df_f)
                    st.session_state["ra_por_aluno"] = ra_map
                    # alerta se o rodízio atual do grupo não bate com a especialidade
                    _rod = ""
                    if "Rodízio Atual" in df_f.columns and df_f["Rodízio Atual"].notna().any():
                        _rod = str(df_f["Rodízio Atual"].dropna().iloc[0])
                    _esp = _sem_acento(especialidade)
                    if _esp and _rod and _esp not in _sem_acento(_rod):
                        st.warning(
                            f"⚠️ O **{grupo_sel}** está em **{_rod.split('(')[0].strip()}** neste período — "
                            f"isso não bate com a especialidade **{especialidade}**. "
                            f"Escolha o grupo correto no seletor acima."
                        )
                    st.success(f"✅ {grupo_sel}: {len(df_f)} alunos em {len(alunos_por_sg)} SGs (com RA). "
                               f"Primeiros: {', '.join(df_f['Nome Completo'].head(3).astype(str))}…")
        except Exception as e:
            st.error(f"Erro: {e}")

    # Sem base enviada, mas veio de importação: mostra os alunos importados editáveis
    elif pf.get("alunos_por_sg"):
        if pf.get("_alunos_base_grupo"):
            st.success(f"✅ Alunos carregados da base oficial — **{pf['_alunos_base_grupo']}**.")
        st.info("💡 Para trocar de grupo (A/B/C...), suba a **base de alunos** acima — aparecerá o seletor de grupo.")
        st.caption("✏️ Alunos importados — edite se necessário:")
        for sg, nomes in sorted(pf["alunos_por_sg"].items(), key=lambda x: int(x[0])):
            k = f"sg_imp_{sg}"
            if k not in st.session_state:  # semeia com os nomes importados (1ª vez)
                st.session_state[k] = "\n".join(nomes)
            atual = [n.strip() for n in st.session_state[k].split("\n") if n.strip()]
            with st.expander(f"SG{sg} — {len(atual)} alunos", expanded=False):
                txt = st.text_area(f"Alunos SG{sg}", key=k, height=100)
                alunos_por_sg[sg] = [n.strip() for n in txt.strip().split("\n") if n.strip()]

    if not alunos_por_sg:
        st.caption("Digite manualmente:")
        for sg in range(1, int(num_sg)+1):
            txt = st.text_area(f"SG{sg} (um nome por linha)", key=f"sg_{sg}", height=80)
            if txt.strip():
                alunos_por_sg[str(sg)] = [n.strip() for n in txt.strip().split("\n") if n.strip()]

# BLOCO 3 — Locais
# BLOCO 3
with st.expander("⚙️ Bloco 3 — Regras Especiais", expanded=True):
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        regra_quinta = st.text_input("Quinta-feira", value=pf.get("regra_quinta","Sem tarde (ENAMED para todos)"))
        regra_terca = st.text_input("Terça-feira", value=pf.get("regra_terca","Tarde encurtada 12-16h (aula às 16h)"))
        limite_ch = st.number_input("Limite CH máximo (h/sem)", 20, 60, _clamp(pf.get("limite_ch",40), 20, 60, 40))
        limite_min = st.number_input("CH mínima alvo (h/sem)", 0, 60, _clamp(pf.get("limite_min",34), 0, 60, 34),
            help="O sistema completa os turnos até cada aluno chegar perto desta carga (sem passar do máximo).")
    with col_r2:
        limite_abs = st.number_input("Limite CH absoluto (h)", 20, 60, _clamp(pf.get("limite_abs",43), 20, 60, 43))
        regra_fds = st.text_area("Regras de plantão FDS", value=pf.get("regra_fds",""), height=80)
        regras_extras = st.text_area("Outras regras", value=pf.get("regras_extras",""), height=80)
    st.caption("🔗 A opção de **manter o subgrupo no mesmo serviço por N dias seguidos** agora fica "
               "**dentro de cada bloco** (Bloco 5), pra você escolher só nos blocos que quiser.")

# BLOCO 4
with st.expander("📊 Bloco 4 — Formato do Excel", expanded=False):
    abas_excel = st.multiselect("Abas desejadas",
        ["Subgrupos","Calendário de Rodízio","Escala Nominal Detalhada","Resumo de Horas","Escala por Local","Regras e Restrições"],
        default=["Subgrupos","Calendário de Rodízio","Escala Nominal Detalhada","Resumo de Horas"])

# BLOCO 5
with st.expander("📍 Bloco 5 — Blocos de Rodízio", expanded=True):
    pf_locais = pf.get("locais", [])
    num_locais_def = _clamp(pf.get("num_locais", len(pf_locais) if pf_locais else 3), 2, 8, 3)
    st.info(
        "💡 **Bloco ≠ serviço.** Um **bloco** é uma ETAPA do rodízio e pode juntar **1, 2 ou 3 serviços**. "
        "Em cada bloco, use o botão **➕ Adicionar serviço** para agrupar mais de um. Exemplos:\n"
        "- Bloco *Anestesiologia* → 3 serviços (Santa Casa Leme + SCA + Mandic)\n"
        "- Bloco *Enfermaria + Centro Cirúrgico* → 2 serviços\n\n"
        "Num bloco com vários serviços, os alunos **rodam entre eles** ao longo das semanas do bloco."
    )
    num_locais = st.number_input("Número de blocos de rodízio (etapas do rodízio, não serviços)", 2, 8, num_locais_def,
        help="Cada bloco é uma ETAPA do rodízio e pode conter 1, 2 ou 3 serviços (use ➕ Adicionar serviço dentro do bloco).")
    locais = []

    # ── Configuração global do rodízio ──────────────────────────────────────
    st.markdown("---")
    st.markdown("**🔢 Durações dos períodos do rodízio (ex.: 3-3-2)**")
    st.caption(
        f"Estes números são as **durações dos períodos** do rodízio (1º período, 2º, 3º...), "
        f"e a soma deve dar o total de semanas ({int(num_semanas)}). "
        f"⚠️ Não é fixo por bloco: cada SG passa por TODOS os blocos, e quantas semanas ele fica "
        f"em cada um depende de QUANDO chega lá (ex.: quem começa no último bloco faz o 1º período de "
        f"3 semanas ali). O calendário define as semanas exatas de cada SG."
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
        nome_bloco_j = st.session_state.get(f"nome_bloco_{j}") or (pf_locais[j].get("nome_bloco", pf_locais[j].get("nome", f"Bloco {j+1}")) if j < len(pf_locais) else f"Bloco {j+1}")
        with cols_rot[j]:
            val = st.number_input(
                f"Bloco {j+1}",
                min_value=1, max_value=int(num_semanas),
                value=st.session_state[key_rot][j],
                key=f"rot_sem_{j}",
                help=nome_bloco_j
            )
            rotacao_semanas.append(val)
            st.caption(f"🏥 {nome_bloco_j}")

    soma_rot = sum(rotacao_semanas)
    if soma_rot == int(num_semanas):
        st.success(f"✅ {' + '.join(str(v) for v in rotacao_semanas)} = {soma_rot} semanas — OK!")
        st.session_state[key_rot] = rotacao_semanas
    elif soma_rot < int(num_semanas):
        st.warning(f"⚠️ Soma atual: {soma_rot} de {int(num_semanas)} semanas — faltam {int(num_semanas)-soma_rot}")
    else:
        st.error(f"❌ Soma atual: {soma_rot} — excede {int(num_semanas)} semanas em {soma_rot-int(num_semanas)}")

    st.markdown("---")

    def _servico_form(key_prefix, pl_srv, label, expanded=True, outros_servicos=None):
        """Renderiza formulário de um serviço dentro de um bloco."""
        with st.expander(f"⚙️ {label}", expanded=expanded):
            ca, cb, cc = st.columns(3)
            with ca:
                _serv_cad = st.session_state.get("_serv_opts_cad", [])
                _nome_cur = pl_srv.get("nome", "")
                _SERV_OUTRO = "➕ Outro (digitar)…"
                if _serv_cad:
                    _nome_opts = list(_serv_cad)
                    if _nome_cur and _nome_cur not in _nome_opts:
                        _nome_opts = [_nome_cur] + _nome_opts
                    _nome_opts = _nome_opts + [_SERV_OUTRO]
                    _nome_sel = st.selectbox("Nome", _nome_opts,
                        index=(_nome_opts.index(_nome_cur) if _nome_cur in _nome_opts else 0),
                        key=f"{key_prefix}_nome_sel",
                        help="Serviço/local cadastrado no Lovable. Se faltar, use 'Outro (digitar)'.")
                    if _nome_sel == _SERV_OUTRO:
                        nome = st.text_input("Nome (digitar)", value=_nome_cur, key=f"{key_prefix}_nome", placeholder="ex: Enfermaria")
                    else:
                        nome = _nome_sel
                else:
                    # catálogo indisponível -> mantém digitação livre (comportamento antigo)
                    nome = st.text_input("Nome", value=_nome_cur, key=f"{key_prefix}_nome", placeholder="ex: Enfermaria")
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
                    min_m = st.number_input("Mín/dia", 0, 20, _clamp(pl_srv.get("min_manha",0), 0, 20, 0), key=f"{key_prefix}_mnm")
                    max_m = st.number_input("Máx/dia", 0, 20, _clamp(pl_srv.get("max_manha",6), 0, 20, 6), key=f"{key_prefix}_mxm")
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
                    min_t = st.number_input("Mín/dia", 0, 20, _clamp(pl_srv.get("min_tarde",3), 0, 20, 3), key=f"{key_prefix}_mnt")
                    max_t = st.number_input("Máx/dia", 0, 20, _clamp(pl_srv.get("max_tarde",4), 0, 20, 4), key=f"{key_prefix}_mxt")
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
                    min_c = st.number_input("Mín/dia", 0, 10, _clamp(pl_srv.get("min_cind",0), 0, 10, 0), key=f"{key_prefix}_mnc")
                    max_c = st.number_input("Máx/dia", 0, 10, _clamp(pl_srv.get("max_cind",2), 0, 10, 2), key=f"{key_prefix}_mxc")
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
                min_fm = st.number_input("Mín", 0,10,_clamp(pl_srv.get("fds_min_manha",1),0,10,1), key=f"{key_prefix}_mnfm") if tem_fm else 0
                max_fm = st.number_input("Máx", 0,10,_clamp(pl_srv.get("fds_max_manha",2),0,10,2), key=f"{key_prefix}_mxfm") if tem_fm else 0
            with tf2:
                st.markdown("**🌇 Tarde FDS**")
                tem_ft = st.checkbox("Tem?", value=bool(pl_srv.get("fds_tarde","")), key=f"{key_prefix}_tft")
                hor_ft = st.text_input("Horário", value=pl_srv.get("fds_tarde","13-19h"), key=f"{key_prefix}_hft") if tem_ft else ""
                min_ft = st.number_input("Mín", 0,10,_clamp(pl_srv.get("fds_min_tarde",1),0,10,1), key=f"{key_prefix}_mnft") if tem_ft else 0
                max_ft = st.number_input("Máx", 0,10,_clamp(pl_srv.get("fds_max_tarde",2),0,10,2), key=f"{key_prefix}_mxft") if tem_ft else 0
            with tf3:
                st.markdown("**🌙 Cinderela FDS**")
                tem_fc = st.checkbox("Tem?", value=bool(pl_srv.get("fds_cind","")), key=f"{key_prefix}_tfc")
                hor_fc = st.text_input("Horário", value=pl_srv.get("fds_cind","19-23h"), key=f"{key_prefix}_hfc") if tem_fc else ""
                min_fc = st.number_input("Mín", 0,10,_clamp(pl_srv.get("fds_min_cind",0),0,10,0), key=f"{key_prefix}_mnfc") if tem_fc else 0
                max_fc = st.number_input("Máx", 0,10,_clamp(pl_srv.get("fds_max_cind",2),0,10,2), key=f"{key_prefix}_mxfc") if tem_fc else 0

            quem_fds = st.text_input("Quem faz FDS?", value=pl_srv.get("fds_quem",""), key=f"{key_prefix}_fdsquem")
            comp_fds = st.text_input("Compensação FDS?", value=pl_srv.get("fds_comp",""), key=f"{key_prefix}_fdscomp")

            st.markdown("---")
            _opc_prio = {"Não priorizar": "", "Manhã": "manha", "Tarde": "tarde"}
            _prio_atual = pl_srv.get("priorizar_periodo", "manha" if pl_srv.get("priorizar_manha") else "")
            _prio_label = next((k for k, v in _opc_prio.items() if v == _prio_atual), "Não priorizar")
            prioriza_label = st.selectbox(
                "🎯 Priorizar um período neste serviço",
                list(_opc_prio.keys()),
                index=list(_opc_prio.keys()).index(_prio_label),
                key=f"{key_prefix}_prioper",
                help="Escolha um período como prioritário. O OUTRO período entra SÓ para quem fez o "
                     "prioritário no mesmo dia/serviço; quem não fez fica com esse período livre (Área Verde). "
                     "Ex.: 'Manhã' → a tarde só para quem foi de manhã. 'Tarde' → a manhã só para quem vai à tarde. "
                     "Pode reduzir a carga horária de alguns alunos — use quando um período é o foco.")
            priorizar_periodo = _opc_prio[prioriza_label]

            # Plantão exclusivo: só alunos que, na mesma semana, estão em OUTRO serviço do bloco
            _excl_atual = pl_srv.get("exclusivo_de_servico", "")
            _excl_on = st.checkbox(
                "🔒 Plantão exclusivo de quem está em OUTRO serviço nesta semana",
                value=bool(_excl_atual), key=f"{key_prefix}_exclon",
                help="Marque para que ESTE serviço/plantão só receba alunos que, na MESMA semana, "
                     "estão escalados em outro serviço do MESMO bloco (ex.: a Enfermaria). "
                     "Útil p/ plantões que devem ser cobertos por quem já está na enfermaria. "
                     "Modo rígido: se não houver alunos elegíveis, o turno pode ficar sem cobertura.")
            exclusivo_de_servico = ""
            exclusivo_turnos = []
            if _excl_on:
                _opts = [o for o in (outros_servicos or []) if o]
                if _opts:
                    _def_idx = 0
                    if _excl_atual:
                        _ea = _sem_acento(_excl_atual)
                        _m = next((ii for ii, o in enumerate(_opts)
                                   if _sem_acento(o) == _ea or _ea in _sem_acento(o) or _sem_acento(o) in _ea), None)
                        if _m is not None:
                            _def_idx = _m
                    exclusivo_de_servico = st.selectbox(
                        "Serviço de origem (escolha um serviço deste bloco)",
                        _opts, index=_def_idx, key=f"{key_prefix}_exclserv_sel",
                        help="O aluno precisa estar NESTE serviço, na mesma semana, para poder pegar este plantão.")
                else:
                    exclusivo_de_servico = st.text_input(
                        "Serviço de origem (nome ou abreviação — ex.: Enfermaria)",
                        value=_excl_atual, key=f"{key_prefix}_exclserv",
                        help="Adicione os outros serviços do bloco (botão ➕ abaixo) para escolher na lista. "
                             "Deve ser um serviço do MESMO bloco.")
                st.caption("A quais turnos/plantões deste serviço a exclusividade se aplica? "
                           "(marque um, vários ou todos)")
                _exc_atual = pl_srv.get("exclusivo_turnos")
                if _exc_atual is None:   # padrão: todos (mantém compatibilidade)
                    _exc_atual = ["du_manha", "du_tarde", "du_cind", "fds_manha", "fds_tarde", "fds_cind"]
                _exa, _exb = st.columns(2)
                with _exa:
                    st.markdown("📅 **Dia útil**")
                    for _k, _lbl in [("du_manha", "Manhã"), ("du_tarde", "Tarde"), ("du_cind", "Cinderela")]:
                        if st.checkbox(_lbl, value=(_k in _exc_atual), key=f"{key_prefix}_exct_{_k}"):
                            exclusivo_turnos.append(_k)
                with _exb:
                    st.markdown("🏖️ **Fim de semana (FDS)**")
                    for _k, _lbl in [("fds_manha", "Manhã"), ("fds_tarde", "Tarde"), ("fds_cind", "Cinderela")]:
                        if st.checkbox(_lbl, value=(_k in _exc_atual), key=f"{key_prefix}_exct_{_k}"):
                            exclusivo_turnos.append(_k)
                if not exclusivo_turnos:
                    st.warning("Selecione ao menos um turno — senão a exclusividade vale para TODOS os turnos deste serviço.")

            return {
                "priorizar_periodo": priorizar_periodo,
                "exclusivo_de_servico": exclusivo_de_servico,
                "exclusivo_turnos": exclusivo_turnos,
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
                _nsrv_bloco = 1 + st.session_state.get(f"n_srv_{i}", 0)
                _nomes_srv = [str(st.session_state.get(f"b{i}_s{k}_nome", "") or "").strip() for k in range(_nsrv_bloco)]
                _nomes_srv = [n for n in _nomes_srv if n]
                if _nomes_srv:
                    st.caption(f"🧩 **{len(_nomes_srv)} serviço(s) neste bloco:** " + "  +  ".join(_nomes_srv))
                else:
                    st.caption(f"🧩 {_nsrv_bloco} serviço(s) neste bloco — use ➕ abaixo para adicionar mais")

            # Serviço principal
            n_srv_preview = 1 + st.session_state.get(f"n_srv_{i}", 0)
            # nomes de todos os serviços do bloco (p/ seletor "serviço de origem" do plantão exclusivo)
            _all_nomes_blk = [str(st.session_state.get(f"b{i}_s{k}_nome", "") or "").strip()
                              for k in range(n_srv_preview)]
            def _irmaos(cur):
                return [nm for k, nm in enumerate(_all_nomes_blk) if k != cur and nm]
            label_s0 = f"Serviço 1 de {nome_bloco}" if nome_bloco else f"Serviço 1 do Bloco {i+1}"
            srv_principal = _servico_form(f"b{i}_s0", pl, label_s0, expanded=True,
                                          outros_servicos=_irmaos(0))

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
                srv_extra = _servico_form(f"b{i}_s{j+1}", pl_extra, label_sj, expanded=False,
                                          outros_servicos=_irmaos(j+1))
                srv_extras.append(srv_extra)
                if st.button(f"❌ Remover Serviço {j+2}", key=f"rm_srv_{i}_{j}"):
                    st.session_state[key_n_srv] -= 1; st.rerun()

            if st.button(f"➕ Adicionar outro serviço a este bloco ({nome_bloco or f'Bloco {i+1}'})",
                         key=f"add_srv_{i}", type="primary", use_container_width=True,
                         help="Agrupe 2 ou 3 serviços no MESMO bloco (ex.: as 3 Anestesiologias, ou Enfermaria + Centro Cirúrgico)."):
                st.session_state[key_n_srv] += 1; st.rerun()

            # Duração e distribuição de SGs no nível do BLOCO
            st.markdown("**📅 Distribuição de SGs neste bloco:**")
            if 1 + st.session_state.get(key_n_srv, 0) > 1:
                st.caption("ℹ️ Este bloco tem **mais de um serviço** → aqui você define quantas semanas cada SG "
                           "passa, fazendo o **rodízio interno** entre os serviços do bloco (ex.: ~1 semana em cada).")
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
                    value=_clamp(pl.get("sem_por_sg", default_sem_sg), 1, int(num_semanas), default_sem_sg),
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

            # Continuidade NESTE bloco (só faz sentido com 2+ serviços)
            dias_consec_bloco = 0
            consec_servico_bloco = ""
            if n_srv_bloco > 1:
                # nomes dos serviços do bloco (principal + extras) para escolher onde aplicar a trava
                nomes_srv_bloco = [srv_principal.get("nome") or "Serviço 1"]
                nomes_srv_bloco += [(s.get("nome") or f"Serviço {j+2}") for j, s in enumerate(srv_extras)]
                cc1, cc2, cc3 = st.columns([3, 1, 2])
                with cc1:
                    manter_b = st.checkbox(
                        "🔗 Manter o MESMO aluno no mesmo serviço por vários dias seguidos",
                        value=bool(pl.get("dias_consec", 0)), key=f"consec_chk_{i}",
                        help="Ex: o aluno fica 3 dias seguidos na Enfermaria (de manhã) e só depois roda. "
                             "Ele continua livre para o PA à tarde / cinderela nesses mesmos dias — a trava vale só para o serviço escolhido.")
                with cc2:
                    if manter_b:
                        dias_consec_bloco = int(st.number_input("Dias seguidos", 2, 5,
                            _clamp(pl.get("dias_consec", 3) or 3, 2, 5, 3), key=f"consec_n_{i}"))
                with cc3:
                    if manter_b:
                        idx_def = 0
                        prev_srv = pl.get("consec_servico", "")
                        if prev_srv in nomes_srv_bloco:
                            idx_def = nomes_srv_bloco.index(prev_srv)
                        consec_servico_bloco = st.selectbox(
                            "Em qual serviço?", nomes_srv_bloco, index=idx_def, key=f"consec_srv_{i}",
                            help="Serviço onde os mesmos alunos ficam fixos pelos dias seguidos (ex: Enfermaria).")

            # 🩺 Complemento de CH: plantões em OUTRO serviço (ex.: Ambulatório → plantão no PA)
            comp_ativo = st.checkbox(
                "🩺 Completar a CH deste bloco com PLANTÕES em outro serviço",
                value=bool(pl.get("complemento_ativo")), key=f"comp_chk_{i}",
                help="Para blocos que sozinhos não fecham a CH (ex: Ambulatório). Os alunos deste bloco "
                     "ganham plantões em outro serviço (ex: PA) até chegar perto da CH alvo — sempre "
                     "respeitando o máximo de alunos por turno/dia do serviço de destino.")
            comp_modo_bloco = "auto"
            comp_servico_bloco = ""
            comp_periodos = None
            comp_obrigatorio = False
            comp_compensa = False
            if comp_ativo:
                # serviços de OUTROS blocos já configurados (acima deste)
                servicos_outros = []
                for lc_prev in locais:
                    for s_prev in [lc_prev] + (lc_prev.get("servicos_extras") or []):
                        nm_prev = s_prev.get("nome")
                        if nm_prev and nm_prev not in servicos_outros:
                            servicos_outros.append(nm_prev)
                escolha = st.radio(
                    "Como escolher o serviço do plantão?",
                    ["Deixar a IA definir (serviço com vaga)", "Definir manualmente"],
                    index=(1 if pl.get("complemento_servico") else 0),
                    key=f"comp_modo_{i}", horizontal=True)
                if escolha.startswith("Definir"):
                    comp_modo_bloco = "manual"
                    if servicos_outros:
                        idx_cs = servicos_outros.index(pl["complemento_servico"]) if pl.get("complemento_servico") in servicos_outros else 0
                        comp_servico_bloco = st.selectbox(
                            "Serviço do plantão (de outro bloco)", servicos_outros, index=idx_cs,
                            key=f"comp_serv_{i}")
                    else:
                        comp_servico_bloco = st.text_input(
                            "Serviço do plantão (digite o nome exato, ex: PA mandic)",
                            value=pl.get("complemento_servico",""), key=f"comp_serv_txt_{i}")
                        st.caption("💡 Configure o bloco do PA ANTES deste para escolher na lista.")
                # Em quais períodos o plantão de complemento pode entrar
                st.markdown("**Em quais períodos o plantão de complemento pode entrar?**")
                st.caption("Marque um, vários ou todos. A IA escolhe o melhor dia dentro do que você marcar "
                           "(considerando quantos alunos já estão naquele dia/turno).")
                _per_atual = pl.get("complemento_periodos")
                if _per_atual is None:
                    _per_atual = ["du_manha", "du_tarde", "du_cind", "fds_manha", "fds_tarde", "fds_cind"]
                comp_periodos = []
                _cda, _cdb = st.columns(2)
                with _cda:
                    st.markdown("📅 **Dia útil**")
                    for _k, _lbl in [("du_manha", "Manhã"), ("du_tarde", "Tarde"), ("du_cind", "Cinderela")]:
                        if st.checkbox(_lbl, value=(_k in _per_atual), key=f"comp_per_{i}_{_k}"):
                            comp_periodos.append(_k)
                with _cdb:
                    st.markdown("🏖️ **Fim de semana**")
                    for _k, _lbl in [("fds_manha", "Manhã"), ("fds_tarde", "Tarde"), ("fds_cind", "Cinderela")]:
                        if st.checkbox(_lbl, value=(_k in _per_atual), key=f"comp_per_{i}_{_k}"):
                            comp_periodos.append(_k)
                if not comp_periodos:
                    st.warning("Selecione ao menos um período (senão o complemento usará todos por padrão).")
                comp_obrigatorio = st.checkbox(
                    "📌 Plantão OBRIGATÓRIO — cada aluno do bloco faz ≥1 por semana (mesmo já tendo batido a CH)",
                    value=bool(pl.get("complemento_obrigatorio")), key=f"comp_obrig_{i}",
                    help="Marque para que TODO aluno deste bloco faça pelo menos 1 plantão por semana nos "
                         "períodos marcados acima — mesmo que já tenha batido a carga horária. "
                         "Ex.: Ambulatório que faz Plantão PA no FDS toda semana. Respeita o máximo de "
                         "alunos por turno do serviço de destino e o limite absoluto de CH.")
                if comp_obrigatorio:
                    comp_compensa = st.checkbox(
                        "↩️ Compensar: tirar 1 período de DIA ÚTIL de quem fez o plantão (respeitando o mínimo do serviço)",
                        value=bool(pl.get("complemento_compensa")), key=f"comp_comp_{i}",
                        help="Quem faz o plantão obrigatório tem 1 turno de dia útil retirado (preferindo a TARDE) "
                             "para compensar as horas — só se o serviço continuar com o mínimo de alunos naquele "
                             "dia/turno E o aluno continuar com a CH mínima. O período liberado vira Área Verde.")
                st.caption("⚠️ O sistema **respeita o máximo de alunos por turno/dia** do serviço de destino "
                           "e **avisa** (na validação) se não houver vaga suficiente para todos atingirem a CH.")

            # Gerar duracao_sgs automaticamente
            duracao_sgs_bloco = {str(sg+1): int(sem_por_sg) for sg in range(n_sgs_total)}

            # Propagar para o serviço principal
            srv_principal["duracao_por_sg"] = duracao_sgs_bloco
            srv_principal["sem_por_sg"] = int(sem_por_sg)
            srv_principal["sgs_por_servico"] = sgs_por_srv
            srv_principal["dias_consec"] = dias_consec_bloco
            srv_principal["consec_servico"] = consec_servico_bloco
            srv_principal["complemento_ativo"] = bool(comp_ativo)
            srv_principal["complemento_modo"] = comp_modo_bloco
            srv_principal["complemento_servico"] = comp_servico_bloco
            srv_principal["complemento_periodos"] = comp_periodos
            srv_principal["complemento_obrigatorio"] = bool(comp_obrigatorio)
            srv_principal["complemento_compensa"] = bool(comp_compensa)
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
with st.expander("🔄 Bloco 6 — Tabela de Rodízio", expanded=True):
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

    nomes_loc_atual = [(l.get("nome_bloco") or l.get("nome") or "") for l in locais]
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

# BLOCO 7 — Salvar / Carregar definições desta escala (reaproveitar depois)
with st.expander("💾 Bloco 7 — Salvar / Carregar definições desta escala", expanded=False):
    st.caption("Guarde TODA a configuração desta escala (especialidade, ano, turma, grupo, alunos, "
               "blocos, serviços, rodízio e regras) num arquivo, para reabrir e reaproveitar depois "
               "sem precisar preencher tudo de novo.")

    col_sv, col_ld = st.columns(2)

    # --- SALVAR ---
    with col_sv:
        st.markdown("**⬇️ Salvar definições atuais**")
        definicoes = {
            "_tipo": "definicoes_escala", "_versao": 1,
            "resumo": f"{especialidade or 'Escala'} · {ano_curso} · {turma} · {grupo}".strip(" ·"),
            "especialidade": especialidade, "ano_curso": ano_curso,
            "turma": turma, "grupo": grupo,
            "data_inicio": str(data_inicio), "num_semanas": int(num_semanas),
            "num_sg": int(num_sg), "num_locais": int(num_locais),
            "alunos_por_sg": alunos_por_sg,
            "ra_por_aluno": st.session_state.get("ra_por_aluno", {}),
            "locais": locais,
            "rodizio_desc": rodizio_desc,
            "regra_quinta": regra_quinta, "regra_terca": regra_terca,
            "limite_ch": int(limite_ch), "limite_min": int(limite_min),
            "limite_abs": int(limite_abs), "regra_fds": regra_fds,
            "regras_extras": regras_extras,
        }
        _slug = _re_val.sub(r"[^A-Za-z0-9]+", "_",
                       f"{especialidade}_{ano_curso}_{turma}_{grupo}").strip("_") or "escala"
        st.download_button(
            "💾 Baixar arquivo de definições (.json)",
            data=json.dumps(definicoes, ensure_ascii=False, indent=2).encode("utf-8"),
            file_name=f"definicoes_{_slug}.json",
            mime="application/json",
            use_container_width=True,
            help="Salva todas as configurações desta escala num arquivo para reabrir depois.")

    # --- CARREGAR ---
    with col_ld:
        st.markdown("**⬆️ Carregar definições salvas**")
        up_defs = st.file_uploader("Arquivo de definições (.json)", type=["json"], key="upload_defs")
        if up_defs is not None:
            if st.button("📂 Carregar estas definições", use_container_width=True):
                try:
                    loaded = json.load(up_defs)
                    if not isinstance(loaded, dict) or loaded.get("_tipo") != "definicoes_escala":
                        st.error("Arquivo inválido — não parece um arquivo de definições desta escala.")
                    else:
                        ra_load = loaded.get("ra_por_aluno", {}) or {}
                        # limpa TODO o estado do formulário para o prefill assumir
                        for _k in list(st.session_state.keys()):
                            del st.session_state[_k]
                        st.session_state.prefill = loaded
                        st.session_state.ra_por_aluno = ra_load
                        st.success("✅ Definições carregadas! Atualizando o formulário...")
                        st.rerun()
                except Exception as e:
                    st.error(f"Erro ao ler o arquivo: {e}")
        st.caption("Depois de carregar, confira os blocos e clique em **Gerar Escala**.")

# ── PRÉVIA DE VIABILIDADE (antes de gerar) ───────────────────────────────────
st.divider()
_total_alunos = sum(len(v) for v in alunos_por_sg.values()) if alunos_por_sg else 0
if locais and _total_alunos:
    with st.expander("🔎 Prévia de viabilidade (CH por bloco, antes de gerar)", expanded=True):
        prev = previa_viabilidade(locais, _total_alunos, int(num_locais),
                                  int(limite_min), int(limite_ch), regra_quinta)
        st.caption(f"Estimativa com ~{prev[0]['alunos_bloco'] if prev else 0} alunos por bloco "
                   f"(≈ {_total_alunos} alunos ÷ {int(num_locais)} blocos). Alvo: {int(limite_min)}–{int(limite_ch)}h/aluno.")
        for idx, p in enumerate(prev):
            tem_comp = bool(locais[idx].get("complemento_ativo")) if idx < len(locais) else False
            if p["ch_max"] < int(limite_min):
                if tem_comp:
                    st.info(f"ℹ️ **{p['bloco']}**: sozinho o bloco fecha ~**{p['ch_max']}h/aluno** (alvo {int(limite_min)}h), "
                            f"mas você ativou **plantões de complemento em outro serviço** → a diferença é coberta com "
                            f"as **vagas livres (inclusive de dia útil)** do serviço de destino. Confira na escala gerada.")
                else:
                    st.warning(f"⚠️ **{p['bloco']}**: sozinho o bloco fecha ~**{p['ch_max']}h/aluno** (< alvo {int(limite_min)}h). "
                               f"Para fechar a CH, o ideal é ativar o **complemento (plantões em outro serviço)** — ele aproveita o "
                               f"**período livre de dia útil** de outro serviço (não precisa ser cinderela/FDS). "
                               f"Alternativas: aumentar o **Máx/dia**, ativar a cinderela, ou colocar menos alunos no bloco.")
            elif p["ch_min"] > int(limite_ch):
                st.error(f"❌ **{p['bloco']}**: a cobertura mínima já força ~**{p['ch_min']}h/aluno** (> {int(limite_ch)}h). "
                         f"Reduza o **Mín/dia** dos turnos.")
            else:
                st.success(f"✅ **{p['bloco']}**: dá pra atingir o alvo "
                           f"(~{int(limite_min)}–{int(min(p['ch_max'], limite_ch))}h/aluno).")

# ── GERAR ────────────────────────────────────────────────────────────────────
if st.button("🚀 Gerar Escala com IA", type="primary", use_container_width=True):
    # Subgrupos esperados que estão sem alunos
    sgs_vazios = [str(i) for i in range(1, int(num_sg) + 1)
                  if not alunos_por_sg.get(str(i))]
    if not especialidade or not turma or not grupo:
        st.error("Preencha Especialidade, Turma e Grupo!")
    elif not alunos_por_sg:
        st.error("Adicione os alunos!")
    elif sgs_vazios:
        st.error(f"⚠️ Os subgrupos {', '.join('SG'+s for s in sgs_vazios)} estão SEM alunos. "
                 f"Preencha todos os {int(num_sg)} subgrupos no Bloco 2 antes de gerar "
                 f"(senão a escala fica incompleta e com carga horária errada).")
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
            "ra_por_aluno": st.session_state.get("ra_por_aluno", {}),
            "rodizio_desc": rodizio_desc,
            "regras_especiais": {"quinta": regra_quinta, "terca": regra_terca,
                "limite_ch": int(limite_ch), "limite_abs": int(limite_abs),
                "limite_min": int(limite_min), "fds": regra_fds},
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

            st.caption("Passo 2/2 — Escala detalhada (montada pelo sistema a partir do calendário)")
            with st.spinner("Montando a escala dia a dia... ⏳"):
                det = gerar_detalhada_python(dados1.get("calendario_rodizio", []), st.session_state.config_atual)

            if det:
                dados1["escala_detalhada"] = det
                dados1["resumo_horas"] = recalcular_resumo_horas(dados1, st.session_state.config_atual)
                st.success(f"✅ Escala detalhada: {len(det)} entradas — montada instantaneamente, respeitando as regras.")
            else:
                st.warning("⚠️ Não consegui montar a escala detalhada — verifique se o calendário de rodízio foi gerado.")

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

    # Detecta semanas faltantes (geração parcial por rate limit) além do caso totalmente vazio
    cfg_atual = st.session_state.get("config_atual", {})
    n_sem_atual = int(cfg_atual.get("num_semanas", 8))
    det_atual = dados_atual.get("escala_detalhada") or []
    semanas_presentes = {int(e.get("semana")) for e in det_atual
                         if str(e.get("semana", "")).isdigit()}
    semanas_faltantes = [s for s in range(1, n_sem_atual + 1) if s not in semanas_presentes]

    if not det_atual or semanas_faltantes:
        if not det_atual:
            st.warning("⚠️ Escala detalhada vazia — as abas Subgrupo e Individual não terão dados.")
        else:
            st.warning(f"⚠️ Faltam as semanas {', '.join(map(str, semanas_faltantes))}.")
        if st.button("🔄 Montar escala detalhada (sistema)", type="primary"):
            with st.spinner("Montando a escala dia a dia... ⏳"):
                novas = gerar_detalhada_python(dados_atual.get("calendario_rodizio", []), cfg_atual)
            if novas:
                dados_atual["escala_detalhada"] = novas
                dados_atual["resumo_horas"] = recalcular_resumo_horas(dados_atual, cfg_atual)
                st.session_state.escala_gerada = json.dumps(dados_atual, ensure_ascii=False)
                st.success(f"✅ {len(novas)} entradas montadas!")
                st.rerun()
            else:
                st.error("Não consegui montar — verifique se o calendário de rodízio existe.")

    mostrar_resultado(
        st.session_state.escala_gerada,
        st.session_state.get("esp_atual",""),
        st.session_state.get("grupo_atual",""),
        st.session_state.get("turma_atual","")
    )

    # ── Avaliar escalação de acordo com pedidos extraordinários ──────────────
    st.divider()
    st.header("🗓️ Avaliar pedidos extraordinários")
    det_pe = dados_atual.get("escala_detalhada") or []
    if not det_pe:
        st.info("Monte a escala detalhada acima para poder avaliar os pedidos extraordinários.")
    else:
        st.caption("Informe os alunos que **não podem** ser escalados em dias específicos. "
                   "O sistema mostra se eles realmente estão escalados nesses dias e, se quiser, "
                   "tenta as trocas — repondo o serviço perdido e respeitando o mínimo de alunos por local.")

        up = st.file_uploader("Carregar pedidos (CSV/XLSX) — precisa de uma coluna **Aluno** e ao menos uma de **Data**",
                              type=["csv", "xlsx"], key="pe_upload")
        if up is not None:
            peds = _pedidos_de_arquivo(up)
            if peds:
                st.session_state.pe_pedidos = peds
                st.success(f"{len(peds)} aluno(s) carregado(s) do arquivo.")
            else:
                st.error("Não reconheci as colunas. Preciso de uma coluna de **Aluno** e ao menos uma de **Data**.")

        base = st.session_state.get("pe_pedidos")
        if base:
            df_ini = pd.DataFrame([{"Aluno": p["aluno"], "Datas": "; ".join(p["datas"])} for p in base])
        else:
            df_ini = pd.DataFrame([{"Aluno": "", "Datas": ""} for _ in range(5)])
        st.markdown("**Ou edite manualmente** — datas separadas por `;` (ex.: `14/08; 15/08`):")
        # chave varia com o conteúdo carregado para o editor recarregar após um novo upload
        sig = "|".join(f'{p["aluno"]}:{",".join(p["datas"])}' for p in (base or []))
        df_edit = st.data_editor(df_ini, num_rows="dynamic", use_container_width=True,
                                 key=f"pe_editor_{abs(hash(sig))}")

        c1, c2 = st.columns(2)
        if c1.button("🔍 Avaliar escalação", type="primary", key="pe_avaliar"):
            pedidos = _parse_pedidos_editor(df_edit)
            if not pedidos:
                st.warning("Informe ao menos um aluno com data(s).")
            else:
                st.session_state.pe_pedidos = pedidos
                st.session_state.pe_avaliacao = avaliar_pedidos_extraordinarios(dados_atual, cfg_atual, pedidos)
                st.session_state.pop("pe_relatorio", None)

        avaliacao = st.session_state.get("pe_avaliacao")
        if avaliacao:
            def _situacao(r):
                if not r["no_periodo"]:
                    return "— fora do período da escala"
                if r["aluno_escala"] == "(não encontrado)":
                    return "aluno não está nesta escala"
                return "🔴 ESCALADO" if r["escalado"] else "🟢 livre"
            tabela = [{
                "Aluno": r["aluno_escala"], "Data": r["data"], "Situação": _situacao(r),
                "Onde está escalado": "; ".join(f'{e["local"]} / {e["turno"]} {e["horario"]}' for e in r["entradas"]),
            } for r in avaliacao]
            st.dataframe(pd.DataFrame(tabela), use_container_width=True, hide_index=True)

            n_conf = sum(len(r["entradas"]) for r in avaliacao if r["escalado"])
            if n_conf:
                st.warning(f"⚠️ {n_conf} conflito(s): aluno escalado num dia em que pediu dispensa.")
                if c2.button("🤖 Tentar trocas automaticamente", key="pe_trocar"):
                    with st.spinner("Aplicando dispensas, substitutos e reposições..."):
                        novo, relat = aplicar_trocas_pedidos(dados_atual, cfg_atual, avaliacao)
                    st.session_state.escala_gerada = json.dumps(novo, ensure_ascii=False)
                    st.session_state.pe_relatorio = relat
                    st.session_state.pe_avaliacao = avaliar_pedidos_extraordinarios(
                        novo, cfg_atual, st.session_state.get("pe_pedidos", []))
                    st.rerun()
            else:
                st.success("✅ Nenhum dos alunos avaliados está escalado nos dias pedidos.")

        relat = st.session_state.get("pe_relatorio")
        if relat:
            st.markdown("#### 🔄 Trocas aplicadas")
            st.dataframe(pd.DataFrame([{
                "Aluno": r["aluno"], "Dia dispensado": f'{r["data"]} ({r["dia"]})', "Serviço": r["servico"],
                "Substituto": r["substituto"], "Reposição": r["reposicao"],
                "Cobertura": r["cobertura"], "Status": r["status"],
            } for r in relat]), use_container_width=True, hide_index=True)
            try:
                val_pe = validar_escala(json.loads(st.session_state.escala_gerada), cfg_atual)
                if val_pe.get("ok"):
                    st.success("✅ Escala revalidada após as trocas: sem violações de cobertura/CH.")
                else:
                    st.warning("⚠️ Após as trocas restam pontos a revisar — confira a Validação no topo do Resultado.")
            except Exception:
                pass
            st.caption("As trocas já estão aplicadas na escala atual — reexporte na seção 📥 Exportar acima.")

# ════════════════════════════════════════════════════════════════════════════