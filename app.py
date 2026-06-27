import streamlit as st
import pandas as pd
import json
import datetime
import requests
import csv
import io

st.set_page_config(page_title="Gerador de Escalas Médicas", page_icon="🏥", layout="wide")

# ── API ──────────────────────────────────────────────────────────────────────
def chamar_claude(mensagens, system_prompt="", max_tokens=8000):
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        st.error("Chave de API não configurada! Vá em Settings → Secrets.")
        return None
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={"x-api-key": api_key, "anthropic-version": "2023-06-01", "content-type": "application/json"},
        json={"model": "claude-sonnet-4-6", "max_tokens": max_tokens,
              "system": system_prompt, "messages": mensagens},
        timeout=180
    )
    result = resp.json()
    if "content" not in result:
        st.error(f"Erro API: {result}")
        return None
    return result["content"][0]["text"]

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

Estrutura obrigatória:
{
  "confirmacao": "resumo em 2 frases",
  "calendario_rodizio": [
    {"semana": 1, "periodo": "06/07-12/07", "alocacao": {"SG1": "Enf", "SG2": "PS"}}
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

Formato compacto — 1 linha por (semana+data+local+turno), listando todos os alunos do SG:
{
  "escala_detalhada": [
    {"semana": 1, "data": "06/07", "dia": "Seg", "local": "Enf", "turno": "Manhã", "horario": "07-13h", "horas": 6, "sg": 1, "alunos": ["Nome1","Nome2","Nome3"]}
  ]
}"""

# ── Mostrar resultado ────────────────────────────────────────────────────────
def mostrar_resultado(resposta_raw, esp, grupo, turma):
    dados = extrair_json(resposta_raw)
    if dados is None:
        st.subheader("📄 Resposta da IA")
        st.text_area("Conteúdo gerado", resposta_raw, height=300)
        st.warning("A IA respondeu em formato inesperado. Verifique o conteúdo acima.")
        return

    # Confirmação
    if dados.get("confirmacao"):
        with st.expander("📋 O que a IA entendeu", expanded=False):
            st.write(dados["confirmacao"])

    # Auditoria
    audit = dados.get("auditoria", {})
    if audit.get("aprovado"):
        st.success("✅ Auditoria aprovada!")
    for err in audit.get("erros", []):
        st.error(f"❌ {err}")
    for av in audit.get("avisos", []):
        st.warning(f"⚠️ {av}")

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
                xls = pd.ExcelFile(arquivo_escala)
                conteudo = ""
                for sheet in xls.sheet_names:
                    df = pd.read_excel(arquivo_escala, sheet_name=sheet, header=None)
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
                        xls_a = pd.ExcelFile(arquivo_alunos)
                        sheets_g = [s for s in xls_a.sheet_names if "GRUPO" in s.upper()]
                        gp = pf.get("grupo","").upper().replace("GRUPO","").strip()
                        sh = next((s for s in sheets_g if gp in s.upper()), sheets_g[0] if sheets_g else None)
                        if sh:
                            df_r = pd.read_excel(arquivo_alunos, sheet_name=sh)
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
            xls2 = pd.ExcelFile(arquivo_alunos)
            grupos_disp = [s for s in xls2.sheet_names if "GRUPO" in s.upper()]
            grupo_sel = st.selectbox("Selecione o grupo", grupos_disp) if grupos_disp else None
            if grupo_sel:
                df_g = pd.read_excel(arquivo_alunos, sheet_name=grupo_sel)
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
with st.expander("📍 Bloco 3 — Locais de Rodízio", expanded=True):
    pf_locais = pf.get("locais", [])
    num_locais_def = int(pf.get("num_locais", len(pf_locais) if pf_locais else 3))
    num_locais = st.number_input("Número de locais", 2, 8, num_locais_def)
    locais = []
    for i in range(int(num_locais)):
        # Pegar dados pré-preenchidos para este local
        pl = pf_locais[i] if i < len(pf_locais) else {}
        with st.container():
            st.markdown(f"### Local {i+1}")
            col_a, col_b = st.columns(2)
            with col_a:
                nome_l = st.text_input("Nome do local", value=pl.get("nome",""), key=f"ln_{i}", placeholder="ex: Enfermaria")
                abrev_l = st.text_input("Abreviação", value=pl.get("abrev",""), key=f"la_{i}", placeholder="ex: Enf")
            with col_b:
                obs_l = st.text_input("Observações", value=pl.get("obs",""), key=f"lobs_{i}", placeholder="ex: FDS feito pelos alunos do Amb")
                unir = st.checkbox("🔗 Unir com outro serviço", value=bool(pl.get("servico2","")), key=f"lunir_{i}")
                s2_nome = st.text_input("2º serviço", value=pl.get("servico2",""), key=f"ln2_{i}") if unir else ""
                s2_obs = st.text_input("2º serviço (regras)", value=pl.get("servico2_obs",""), key=f"lobs2_{i}") if unir else ""

            # Duração por SG
            n_sgs_atual = len(alunos_por_sg) if alunos_por_sg else int(num_sg)
            st.markdown("**📅 Duração por SG neste local (semanas, 0 = não passa aqui):**")
            duracao_sgs = {}
            dur_cols = st.columns(min(n_sgs_atual, 6))
            pl_dur = pl.get("duracao_por_sg", {})
            for sg_idx in range(n_sgs_atual):
                sg_num = sg_idx + 1
                with dur_cols[sg_idx % 6]:
                    dur = st.number_input(f"SG{sg_num}", 0, int(num_semanas),
                        int(pl_dur.get(str(sg_num), 0)), key=f"dur_{i}_{sg_num}")
                    if dur > 0: duracao_sgs[str(sg_num)] = int(dur)

            st.markdown("**⏰ Turnos:**")
            tab_util, tab_fds = st.tabs(["📅 Dias Úteis", "🏖️ Final de Semana"])

            with tab_util:
                col_u1, col_u2, col_u3 = st.columns(3)
                with col_u1:
                    st.markdown("**🌅 Manhã**")
                    tem_m = st.checkbox("Tem manhã?", value=bool(pl.get("manha","07-13h")), key=f"tm_{i}")
                    hor_m = st.text_input("Horário", value=pl.get("manha","07-13h"), key=f"hm_{i}") if tem_m else ""
                    min_m = st.number_input("Mín/dia", 0, 20, int(pl.get("min_manha",0)), key=f"mnm_{i}") if tem_m else 0
                    max_m = st.number_input("Máx/dia", 0, 20, int(pl.get("max_manha",6)), key=f"mxm_{i}") if tem_m else 0

                with col_u2:
                    st.markdown("**🌇 Tarde**")
                    tem_t = st.checkbox("Tem tarde?", value=bool(pl.get("tarde","12-18h")), key=f"tt_{i}")
                    if tem_t:
                        hor_t = st.text_input("Horário normal", value=pl.get("tarde","12-18h"), key=f"ht_{i}")
                        min_t = st.number_input("Mín/dia", 0, 20, int(pl.get("min_tarde",3)), key=f"mnt_{i}")
                        max_t = st.number_input("Máx/dia", 0, 20, int(pl.get("max_tarde",4)), key=f"mxt_{i}")
                        st.markdown("**🚫 Bloqueios:**")
                        key_bloqs = f"bloqs_{i}"
                        if key_bloqs not in st.session_state:
                            st.session_state[key_bloqs] = pl.get("bloqueios_tarde", [
                                {"dia":"Qui","tipo":"Sem tarde","horario":""}
                            ])
                        col_ba, col_bb = st.columns(2)
                        with col_ba:
                            if st.button("➕ Adicionar", key=f"add_b_{i}"):
                                st.session_state[key_bloqs].append({"dia":"Seg","tipo":"Sem tarde","horario":""})
                                st.rerun()
                        with col_bb:
                            if st.session_state[key_bloqs] and st.button("🗑️ Limpar todos", key=f"clr_b_{i}"):
                                st.session_state[key_bloqs] = []
                                st.rerun()
                        if not st.session_state[key_bloqs]:
                            st.caption("Sem bloqueios — tarde em todos os dias úteis.")
                        bloqueios_t = []
                        dias_op = ["Seg","Ter","Qua","Qui","Sex"]
                        for b, bloq in enumerate(st.session_state[key_bloqs]):
                            bc1,bc2,bc3,bc4 = st.columns([2,2,2,1])
                            with bc1:
                                idx_d = dias_op.index(bloq.get("dia","Qui")) if bloq.get("dia") in dias_op else 3
                                d_b = st.selectbox("Dia", dias_op, index=idx_d, key=f"dbloq_{i}_{b}")
                            with bc2:
                                t_b = st.selectbox("Tipo", ["Sem tarde","Horário reduzido"],
                                    index=1 if bloq.get("tipo")=="Horário reduzido" else 0, key=f"tbloq_{i}_{b}")
                            with bc3:
                                h_b = st.text_input("Horário", value=bloq.get("horario","12-16h"), key=f"hbloq_{i}_{b}") if t_b=="Horário reduzido" else ""
                                if t_b != "Horário reduzido": st.text_input("—", disabled=True, key=f"hbloq_{i}_{b}")
                            with bc4:
                                st.write("")
                                if st.button("❌", key=f"del_b_{i}_{b}"):
                                    st.session_state[key_bloqs].pop(b); st.rerun()
                            bloqueios_t.append({"dia":d_b,"tipo":t_b,"horario":h_b})
                    else:
                        hor_t,min_t,max_t,bloqueios_t = "",0,0,[]

                with col_u3:
                    st.markdown("**🌙 Cinderela**")
                    tem_c = st.checkbox("Tem cinderela?", value=bool(pl.get("cinderela","")), key=f"tc_{i}")
                    if tem_c:
                        hor_c = st.text_input("Horário", value=pl.get("cinderela","19-23h"), key=f"hc_{i}")
                        min_c = st.number_input("Mín/dia", 0, 10, int(pl.get("min_cind",0)), key=f"mnc_{i}")
                        max_c = st.number_input("Máx/dia", 0, 10, int(pl.get("max_cind",2)), key=f"mxc_{i}")
                        dias_c_validos = ["Seg","Ter","Qua","Qui","Sex"]
                        dias_c_default = [d for d in pl.get("dias_cind",["Sex"]) if d in dias_c_validos] or ["Sex"]
                        dias_c = st.multiselect("Dias", dias_c_validos,
                            default=dias_c_default, key=f"dc_{i}")
                    else:
                        hor_c,min_c,max_c,dias_c = "",0,0,[]

            with tab_fds:
                tem_fds = st.checkbox("Tem plantão no FDS?", value=pl.get("fds",False), key=f"lfds_{i}")
                if tem_fds:
                    col_f1,col_f2,col_f3 = st.columns(3)
                    with col_f1:
                        st.markdown("**🌅 Manhã FDS**")
                        tem_fm = st.checkbox("Tem?", value=bool(pl.get("fds_manha","")), key=f"tfm_{i}")
                        hor_fm = st.text_input("Horário", value=pl.get("fds_manha","07-12h"), key=f"hfm_{i}") if tem_fm else ""
                        min_fm = st.number_input("Mín", 0, 10, int(pl.get("fds_min_manha",1)), key=f"mnfm_{i}") if tem_fm else 0
                        max_fm = st.number_input("Máx", 0, 10, int(pl.get("fds_max_manha",2)), key=f"mxfm_{i}") if tem_fm else 0
                    with col_f2:
                        st.markdown("**🌇 Tarde FDS**")
                        tem_ft = st.checkbox("Tem?", value=bool(pl.get("fds_tarde","")), key=f"tft_{i}")
                        hor_ft = st.text_input("Horário", value=pl.get("fds_tarde","13-19h"), key=f"hft_{i}") if tem_ft else ""
                        min_ft = st.number_input("Mín", 0, 10, int(pl.get("fds_min_tarde",1)), key=f"mnft_{i}") if tem_ft else 0
                        max_ft = st.number_input("Máx", 0, 10, int(pl.get("fds_max_tarde",2)), key=f"mxft_{i}") if tem_ft else 0
                    with col_f3:
                        st.markdown("**🌙 Cinderela FDS**")
                        tem_fc = st.checkbox("Tem?", value=bool(pl.get("fds_cind","")), key=f"tfc_{i}")
                        hor_fc = st.text_input("Horário", value=pl.get("fds_cind","19-23h"), key=f"hfc_{i}") if tem_fc else ""
                        min_fc = st.number_input("Mín", 0, 10, int(pl.get("fds_min_cind",0)), key=f"mnfc_{i}") if tem_fc else 0
                        max_fc = st.number_input("Máx", 0, 10, int(pl.get("fds_max_cind",2)), key=f"mxfc_{i}") if tem_fc else 0
                    quem_fds = st.text_input("Quem faz?", value=pl.get("fds_quem",""), key=f"lfdsquem_{i}")
                    comp_fds = st.text_input("Compensação?", value=pl.get("fds_comp",""), key=f"lfdscomp_{i}")
                else:
                    hor_fm=hor_ft=hor_fc=quem_fds=comp_fds=""
                    min_fm=max_fm=min_ft=max_ft=min_fc=max_fc=0

            locais.append({
                "nome": nome_l, "abrev": abrev_l, "obs": obs_l,
                "servico2": s2_nome, "servico2_obs": s2_obs,
                "duracao_por_sg": duracao_sgs,
                "manha": hor_m if tem_m else "", "min_manha": int(min_m), "max_manha": int(max_m),
                "tarde": hor_t if tem_t else "", "min_tarde": int(min_t), "max_tarde": int(max_t),
                "bloqueios_tarde": bloqueios_t,
                "cinderela": hor_c if tem_c else "", "min_cind": int(min_c), "max_cind": int(max_c), "dias_cind": dias_c,
                "fds": tem_fds,
                "fds_manha": hor_fm, "fds_min_manha": int(min_fm), "fds_max_manha": int(max_fm),
                "fds_tarde": hor_ft, "fds_min_tarde": int(min_ft), "fds_max_tarde": int(max_ft),
                "fds_cind": hor_fc, "fds_min_cind": int(min_fc), "fds_max_cind": int(max_fc),
                "fds_quem": quem_fds, "fds_comp": comp_fds,
                "cov_tarde": int(min_t) if tem_t else 0,
            })
            st.divider()

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
        briefing = f"""
# BRIEFING DE ESCALA MÉDICA

## BLOCO 1
Especialidade: {especialidade} | Ano: {ano_curso} | Turma: {turma} | Grupo: {grupo}
Início: {data_inicio} | Semanas: {num_semanas}

## BLOCO 2 — ALUNOS
{json.dumps(alunos_por_sg, ensure_ascii=False, indent=2)}

## BLOCO 3 — LOCAIS
{json.dumps(locais, ensure_ascii=False, indent=2)}

## BLOCO 4 — RODÍZIO
{rodizio_desc}

## BLOCO 5 — REGRAS
Quinta: {regra_quinta}
Terça: {regra_terca}
Limite CH: {limite_ch}h | Absoluto: {limite_abs}h
FDS: {regra_fds}
Extras: {regras_extras}

## BLOCO 6 — EXCEL
Abas: {', '.join(abas_excel)}
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
            st.session_state.escala_gerada = resposta
            st.session_state.esp_atual = especialidade
            st.session_state.grupo_atual = grupo
            st.session_state.turma_atual = turma
            cal_gerado = json.dumps(dados1.get("calendario_rodizio",[]), ensure_ascii=False)
            with st.spinner("Passo 2/2 — Escala detalhada dia a dia... ⏳"):
                resp2 = chamar_claude(
                    [{"role": "user", "content": f"Briefing:\n{briefing}\n\nCalendário:\n{cal_gerado}\n\nGere a escala_detalhada."}],
                    system_prompt=SYSTEM_DETALHE, max_tokens=12000
                )
            if resp2:
                dados2 = extrair_json(resp2) or {}
                dados1["escala_detalhada"] = dados2.get("escala_detalhada", [])
                st.session_state.escala_gerada = json.dumps(dados1, ensure_ascii=False)
            st.rerun()

# Mostrar resultado
if "escala_gerada" in st.session_state and "esp_atual" in st.session_state:
    st.divider()
    st.header("📊 Resultado")
    mostrar_resultado(
        st.session_state.escala_gerada,
        st.session_state.get("esp_atual",""),
        st.session_state.get("grupo_atual",""),
        st.session_state.get("turma_atual","")
    )

# ════════════════════════════════════════════════════════════════════════════