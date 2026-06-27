import streamlit as st
import json
import datetime
import requests
import pandas as pd
import io

st.set_page_config(page_title="Gerador de Escalas Médicas", page_icon="🏥", layout="wide")
st.title("🏥 Gerador de Escalas Médicas")
st.caption("Powered by Claude AI — geração inteligente com todas as regras da sua especialidade")

# ── MODO ─────────────────────────────────────────────────────────────────────
modo = st.radio("Como deseja começar?",
    ["✨ Criar nova escala", "📂 Importar escala existente e modificar"],
    horizontal=True)
st.divider()

def chamar_claude(mensagens, system_prompt="", max_tokens=4000):
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        st.error("Chave de API não configurada! Vá em Settings → Secrets.")
        return None
    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={"x-api-key": api_key, "anthropic-version": "2023-06-01", "content-type": "application/json"},
        json={"model": "claude-sonnet-4-6", "max_tokens": max_tokens,
              "system": system_prompt, "messages": mensagens},
        timeout=120
    )
    result = resp.json()
    if "content" not in result:
        st.error(f"Erro API: {result}")
        return None
    return result["content"][0]["text"]

# ════════════════════════════════════════════════════════════════════════════
# MODO IMPORTAR
# ════════════════════════════════════════════════════════════════════════════
if modo == "📂 Importar escala existente e modificar":
    st.header("📂 Importar Escala Existente")
    arquivo_escala = st.file_uploader("Upload da escala existente (.xlsx)", type=["xlsx"])

    if arquivo_escala:
        if st.button("🤖 Analisar escala com IA", type="primary", use_container_width=True):
            with st.spinner("Lendo e interpretando sua escala..."):
                xls = pd.ExcelFile(arquivo_escala)
                conteudo = ""
                for sheet in xls.sheet_names:
                    df = pd.read_excel(arquivo_escala, sheet_name=sheet, header=None)
                    conteudo += f"\n\n=== ABA: {sheet} ===\n{df.fillna('').to_string(index=False, header=False)}"
                    if len(conteudo) > 10000:
                        conteudo = conteudo[:10000] + "\n...[truncado]"
                        break

                resposta = chamar_claude([{
                    "role": "user",
                    "content": f"""Analise esta escala médica e extraia um briefing completo em JSON.

CONTEÚDO:
{conteudo}

Retorne APENAS JSON válido (sem markdown) com:
{{
  "especialidade": "",
  "grupo": "",
  "turma": "",
  "data_inicio": "YYYY-MM-DD",
  "num_semanas": 8,
  "num_sg": 6,
  "alunos_por_sg": {{"1": ["Nome1","Nome2"], "2": ["Nome3"]}},
  "locais": [{{"nome":"","abrev":"","fds":false,"turno_m":"07-13h","turno_t":"12-18h","turno_c":null,"cov_tarde":3,"obs":""}}],
  "rodizio": {{"descricao":"","tabela":{{"SG1":["Local1","Local1","Local2"],"SG2":[]}}}},
  "regras_especiais": {{"quinta":"sem tarde (ENAMED)","terca":"tarde 12-16h","limite_ch":40,"limite_abs":43}},
  "resumo": "descrição em português"
}}"""
                }])

                if resposta:
                    try:
                        txt = resposta.strip()
                        if "```" in txt:
                            txt = txt.split("```")[1]
                            if txt.startswith("json"): txt = txt[4:]
                        dados = json.loads(txt)
                        st.session_state.escala_importada = dados
                        st.success("✅ Escala analisada!")
                    except Exception as e:
                        st.error(f"Erro ao parsear: {e}")
                        st.text(resposta[:500])

        if "escala_importada" in st.session_state:
            dados = st.session_state.escala_importada
            st.info(f"💡 {dados.get('resumo','')}")

            st.subheader("✏️ Revise e edite o que quiser:")
            col1, col2 = st.columns(2)
            with col1:
                esp = st.text_input("Especialidade", value=dados.get("especialidade",""), key="imp_esp")
                grupo = st.text_input("Grupo", value=dados.get("grupo",""), key="imp_grupo")
            with col2:
                turma = st.text_input("Turma", value=dados.get("turma",""), key="imp_turma")
                try: d0 = datetime.date.fromisoformat(dados.get("data_inicio", str(datetime.date.today())))
                except: d0 = datetime.date.today()
                data_ini = st.date_input("Data início", value=d0, key="imp_data")

            nsem = st.number_input("Nº de semanas", 1, 20, int(dados.get("num_semanas",8)), key="imp_nsem")

            st.subheader("Alunos por Subgrupo")
            alunos_por_sg = dados.get("alunos_por_sg", {})
            alunos_editados = {}
            for sg, nomes in sorted(alunos_por_sg.items(), key=lambda x: int(x[0])):
                with st.expander(f"SG{sg} ({len(nomes)} alunos)", expanded=False):
                    txt_alunos = st.text_area(f"SG{sg}", value="\n".join(nomes), key=f"imp_sg_{sg}", height=100)
                    alunos_editados[sg] = [n.strip() for n in txt_alunos.strip().split("\n") if n.strip()]

            modificacoes = st.text_area("O que deseja modificar nesta escala? (descreva livremente)",
                placeholder="Ex: Trocar Luiza do SG2 para SG3 na semana 4\nAdicionar feriado em 15/08\nMudar Ambulatório para só manhã nas semanas 5-8",
                height=100, key="imp_modif")

            if st.button("🚀 Gerar Escala com IA", type="primary", use_container_width=True):
                briefing = f"""
ESPECIALIDADE: {esp} | GRUPO: {grupo} | TURMA: {turma}
DATA INÍCIO: {data_ini} | SEMANAS: {nsem}

ALUNOS POR SUBGRUPO:
{json.dumps(alunos_editados, ensure_ascii=False, indent=2)}

LOCAIS E REGRAS:
{json.dumps(dados.get('locais',[]), ensure_ascii=False, indent=2)}

RODÍZIO:
{json.dumps(dados.get('rodizio',{}), ensure_ascii=False, indent=2)}

REGRAS ESPECIAIS:
{json.dumps(dados.get('regras_especiais',{}), ensure_ascii=False, indent=2)}

MODIFICAÇÕES SOLICITADAS:
{modificacoes if modificacoes else "Nenhuma — manter a lógica original"}
"""
                with st.spinner("IA gerando a escala completa... ⏳ (pode levar até 1 minuto)"):
                    system = """Você é um especialista em escalas de internato médico. 
Gere uma escala completa seguindo as regras do briefing.
Responda em JSON com esta estrutura:
{
  "calendario": [
    {"semana": 1, "periodo": "06/07-12/07", "sg1": "Local", "sg2": "Local", ...}
  ],
  "escala_nominal": [
    {"semana": 1, "data": "06/07", "dia": "Seg", "local": "Enf", "turno": "Manhã", "horario": "07-13h", "horas": 6, "alunos": ["Nome1","Nome2"]}
  ],
  "resumo_horas": [
    {"sg": 1, "nome": "Nome", "total_horas": 38, "semanas": [36,40,38,40,36,40,38,40]}
  ],
  "auditoria": {"erros": [], "avisos": [], "aprovado": true}
}"""
                    resposta = chamar_claude(
                        [{"role": "user", "content": f"Gere a escala completa com base neste briefing:\n\n{briefing}"}],
                        system_prompt=system, max_tokens=8000
                    )
                    if resposta:
                        st.session_state.escala_gerada = resposta
                        st.success("✅ Escala gerada!")

            if "escala_gerada" in st.session_state:
                _mostrar_resultado(st.session_state.escala_gerada, esp, grupo, turma)

    st.stop()

# ════════════════════════════════════════════════════════════════════════════
# MODO CRIAR NOVA ESCALA — FORMULÁRIO DE BRIEFING GUIADO
# ════════════════════════════════════════════════════════════════════════════

st.header("📋 Briefing da Escala")
st.caption("Preencha com atenção — quanto mais detalhado, mais precisa a escala gerada.")

# BLOCO 1 — Identificação
with st.expander("📌 Bloco 1 — Identificação", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        especialidade = st.text_input("Especialidade *", placeholder="ex: Clínica Médica")
        ano_curso = st.selectbox("Ano do curso *", ["3º Ano","4º Ano","5º Ano","6º Ano"], index=1)
        turma = st.text_input("Turma *", placeholder="ex: T6")
    with col2:
        grupo = st.text_input("Grupo *", placeholder="ex: Grupo A")
        data_inicio = st.date_input("Data de início (segunda-feira) *")
        num_semanas = st.number_input("Número de semanas *", 1, 20, 8)

# BLOCO 2 — Alunos
with st.expander("👥 Bloco 2 — Alunos e Subgrupos", expanded=True):
    arquivo_alunos = st.file_uploader("Upload Excel de alunos (opcional)", type=["xlsx"])
    num_sg = st.number_input("Número de subgrupos", 2, 8, 6)

    alunos_por_sg = {}
    if arquivo_alunos:
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
                    st.success(f"✅ {len(df_f)} alunos carregados em {len(alunos_por_sg)} SGs")
        except Exception as e:
            st.error(f"Erro: {e}")

    if not alunos_por_sg:
        st.caption("Ou digite os alunos manualmente:")
        for sg in range(1, int(num_sg)+1):
            txt = st.text_area(f"SG{sg} (um nome por linha)", key=f"sg_{sg}", height=80)
            if txt.strip():
                alunos_por_sg[str(sg)] = [n.strip() for n in txt.strip().split("\n") if n.strip()]

# BLOCO 3 — Locais
with st.expander("📍 Bloco 3 — Locais de Rodízio", expanded=True):
    num_locais = st.number_input("Número de locais", 2, 8, 3)
    locais = []
    for i in range(int(num_locais)):
        with st.container():
            st.markdown(f"### Local {i+1}")

            # Identificação
            col_a, col_b = st.columns(2)
            with col_a:
                nome_l = st.text_input("Nome do local", key=f"ln_{i}", placeholder="ex: Enfermaria")
                abrev_l = st.text_input("Abreviação", key=f"la_{i}", placeholder="ex: Enf")
            with col_b:
                obs_l = st.text_input("Observações", key=f"lobs_{i}", placeholder="ex: FDS feito pelos alunos do Amb")
                # Opção de unir 2 serviços em 1 bloco
                unir_servico = st.checkbox("🔗 Unir com outro serviço neste bloco", key=f"lunir_{i}")
                if unir_servico:
                    servico2_nome = st.text_input("Nome do 2º serviço", key=f"ln2_{i}", placeholder="ex: PA Mandic")
                    servico2_obs = st.text_input("Regra do 2º serviço", key=f"lobs2_{i}",
                        placeholder="ex: PA: máx 2 alunos/turno, M+T+C, FDS")
                else:
                    servico2_nome = ""
                    servico2_obs = ""

            # Tabela de turnos por período
            st.markdown("**⏰ Configuração de turnos e vagas:**")

            tab_util, tab_fds = st.tabs(["📅 Dias Úteis", "🏖️ Final de Semana"])

            with tab_util:
                st.caption("Configure cada turno para os dias úteis (Seg–Sex)")
                col_u1, col_u2, col_u3 = st.columns(3)
                with col_u1:
                    st.markdown("**🌅 Manhã**")
                    tem_manha = st.checkbox("Tem manhã?", value=True, key=f"tm_{i}")
                    if tem_manha:
                        hor_manha = st.text_input("Horário", value="07-13h", key=f"hm_{i}")
                        min_manha = st.number_input("Mín alunos/dia", 0, 20, 0, key=f"mnm_{i}")
                        max_manha = st.number_input("Máx alunos/dia", 0, 20, 6, key=f"mxm_{i}")
                    else:
                        hor_manha, min_manha, max_manha = "", 0, 0

                with col_u2:
                    st.markdown("**🌇 Tarde**")
                    tem_tarde = st.checkbox("Tem tarde?", value=True, key=f"tt_{i}")
                    if tem_tarde:
                        hor_tarde = st.text_input("Horário normal", value="12-18h", key=f"ht_{i}")
                        hor_tarde_ter = st.text_input("Horário terça (se diferente)", value="12-16h", key=f"htt_{i}")
                        hor_tarde_qui = st.text_input("Horário quinta (vazio=sem tarde)", value="", key=f"htq_{i}")
                        min_tarde = st.number_input("Mín alunos/dia", 0, 20, 3, key=f"mnt_{i}")
                        max_tarde = st.number_input("Máx alunos/dia", 0, 20, 4, key=f"mxt_{i}")
                    else:
                        hor_tarde, hor_tarde_ter, hor_tarde_qui = "", "", ""
                        min_tarde, max_tarde = 0, 0

                with col_u3:
                    st.markdown("**🌙 Cinderela**")
                    tem_cind = st.checkbox("Tem cinderela?", key=f"tc_{i}")
                    if tem_cind:
                        hor_cind = st.text_input("Horário", value="19-23h", key=f"hc_{i}")
                        min_cind = st.number_input("Mín alunos/dia", 0, 10, 0, key=f"mnc_{i}")
                        max_cind = st.number_input("Máx alunos/dia", 0, 10, 2, key=f"mxc_{i}")
                        dias_cind = st.multiselect("Dias com cinderela", ["Seg","Ter","Qua","Qui","Sex"],
                                                    default=["Sex"], key=f"dc_{i}")
                    else:
                        hor_cind, min_cind, max_cind, dias_cind = "", 0, 0, []

            with tab_fds:
                tem_fds = st.checkbox("Tem plantão no FDS?", key=f"lfds_{i}")
                if tem_fds:
                    col_f1, col_f2, col_f3 = st.columns(3)
                    with col_f1:
                        st.markdown("**🌅 Manhã FDS**")
                        tem_fds_m = st.checkbox("Tem manhã FDS?", value=True, key=f"tfm_{i}")
                        if tem_fds_m:
                            hor_fds_m = st.text_input("Horário", value="07-12h", key=f"hfm_{i}")
                            min_fds_m = st.number_input("Mín alunos", 0, 10, 1, key=f"mnfm_{i}")
                            max_fds_m = st.number_input("Máx alunos", 0, 10, 2, key=f"mxfm_{i}")
                        else:
                            hor_fds_m, min_fds_m, max_fds_m = "", 0, 0
                    with col_f2:
                        st.markdown("**🌇 Tarde FDS**")
                        tem_fds_t = st.checkbox("Tem tarde FDS?", key=f"tft_{i}")
                        if tem_fds_t:
                            hor_fds_t = st.text_input("Horário", value="13-19h", key=f"hft_{i}")
                            min_fds_t = st.number_input("Mín alunos", 0, 10, 1, key=f"mnft_{i}")
                            max_fds_t = st.number_input("Máx alunos", 0, 10, 2, key=f"mxft_{i}")
                        else:
                            hor_fds_t, min_fds_t, max_fds_t = "", 0, 0
                    with col_f3:
                        st.markdown("**🌙 Cinderela FDS**")
                        tem_fds_c = st.checkbox("Tem cinderela FDS?", key=f"tfc_{i}")
                        if tem_fds_c:
                            hor_fds_c = st.text_input("Horário", value="19-23h", key=f"hfc_{i}")
                            min_fds_c = st.number_input("Mín alunos", 0, 10, 0, key=f"mnfc_{i}")
                            max_fds_c = st.number_input("Máx alunos", 0, 10, 2, key=f"mxfc_{i}")
                        else:
                            hor_fds_c, min_fds_c, max_fds_c = "", 0, 0

                    quem_faz_fds = st.text_input("Quem faz o FDS deste local?",
                        key=f"lfdsquem_{i}", placeholder="ex: alunos do Ambulatório | alunos do próprio local")
                    compensacao_fds = st.text_input("Compensação por FDS?",
                        key=f"lfdscomp_{i}", placeholder="ex: perde 1 tarde do Amb | nenhuma")
                else:
                    tem_fds_m = tem_fds_t = tem_fds_c = False
                    hor_fds_m = hor_fds_t = hor_fds_c = ""
                    min_fds_m = max_fds_m = min_fds_t = max_fds_t = min_fds_c = max_fds_c = 0
                    quem_faz_fds = compensacao_fds = ""

            locais.append({
                "nome": nome_l, "abrev": abrev_l, "obs": obs_l,
                "servico2": servico2_nome, "servico2_obs": servico2_obs,
                # Dias úteis
                "manha": hor_manha if tem_manha else "",
                "min_manha": int(min_manha), "max_manha": int(max_manha),
                "tarde": hor_tarde if tem_tarde else "",
                "tarde_terca": hor_tarde_ter if tem_tarde else "",
                "tarde_quinta": hor_tarde_qui if tem_tarde else "",
                "min_tarde": int(min_tarde), "max_tarde": int(max_tarde),
                "cinderela": hor_cind if tem_cind else "",
                "min_cind": int(min_cind), "max_cind": int(max_cind),
                "dias_cind": dias_cind,
                # FDS
                "fds": tem_fds,
                "fds_manha": hor_fds_m if tem_fds_m else "",
                "fds_min_manha": int(min_fds_m), "fds_max_manha": int(max_fds_m),
                "fds_tarde": hor_fds_t if tem_fds_t else "",
                "fds_min_tarde": int(min_fds_t), "fds_max_tarde": int(max_fds_t),
                "fds_cind": hor_fds_c if tem_fds_c else "",
                "fds_min_cind": int(min_fds_c), "fds_max_cind": int(max_fds_c),
                "fds_quem": quem_faz_fds, "fds_comp": compensacao_fds,
                # Compatibilidade com gerador
                "cov_tarde": int(min_tarde),
            })
            st.divider()


# BLOCO 4 — Rodízio
with st.expander("🔄 Bloco 4 — Tabela de Rodízio", expanded=True):
    st.caption("Descreva ou cole a tabela de rodízio. Ex: Par SG1+SG2 → Enf S1-3, PS S4-5, Amb S6-8")
    rodizio_desc = st.text_area("Tabela de rodízio (descreva livremente ou cole a tabela)",
        height=120, placeholder="""Ex:
Par1 = SG1+SG2: Enf semanas 1-3 → PS semanas 4-5 → Amb semanas 6-8
Par2 = SG3+SG4: PS semanas 1-3 → Amb semanas 4-5 → Enf semanas 6-8
Par3 = SG5+SG6: Amb semanas 1-3 → Enf semanas 4-5 → PS semanas 6-8""")

# BLOCO 5 — Regras especiais
with st.expander("⚙️ Bloco 5 — Regras Especiais", expanded=True):
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        regra_quinta = st.text_input("Quinta-feira", value="Sem tarde (ENAMED para todos)",
                                      placeholder="ex: Sem tarde — ENAMED")
        regra_terca = st.text_input("Terça-feira", value="Tarde encurtada 12-16h (aula às 16h)",
                                     placeholder="ex: Tarde 12-16h por causa de aula")
        limite_ch = st.number_input("Limite CH semanal padrão (h)", 20, 60, 40)
    with col_r2:
        limite_abs = st.number_input("Limite CH absoluto (nunca passar)", 20, 60, 43)
        regra_fds = st.text_area("Regras de plantão FDS", height=80,
            placeholder="ex: FDS PS = alunos do Amb, 3M+3T, compensação: perde 1 tarde Amb")
        regras_extras = st.text_area("Outras regras e exceções", height=80,
            placeholder="ex: Ambulatório: manhã 08-12h (não 07-13h)")

# BLOCO 6 — Excel
with st.expander("📊 Bloco 6 — Formato do Excel", expanded=False):
    abas_excel = st.multiselect("Abas desejadas no Excel",
        ["Subgrupos", "Calendário de Rodízio", "Escala Nominal Detalhada",
         "Resumo de Horas", "Escala por Local", "Regras e Restrições"],
        default=["Subgrupos", "Calendário de Rodízio", "Escala Nominal Detalhada", "Resumo de Horas"])
    cols_extras = st.text_input("Colunas extras no Resumo de Horas",
        placeholder="ex: Plantões Enf FDS, Plantões PS FDS, Cinderelas")

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

## BLOCO 1 — IDENTIFICAÇÃO
Especialidade: {especialidade}
Ano: {ano_curso} | Turma: {turma} | Grupo: {grupo}
Início: {data_inicio} | Semanas: {num_semanas}

## BLOCO 2 — ALUNOS POR SUBGRUPO
{json.dumps(alunos_por_sg, ensure_ascii=False, indent=2)}

## BLOCO 3 — LOCAIS
{json.dumps(locais, ensure_ascii=False, indent=2)}

## BLOCO 4 — RODÍZIO
{rodizio_desc}

## BLOCO 5 — REGRAS ESPECIAIS
Quinta-feira: {regra_quinta}
Terça-feira: {regra_terca}
Limite CH padrão: {limite_ch}h | Limite absoluto: {limite_abs}h
Plantões FDS: {regra_fds}
Outras regras: {regras_extras}

## BLOCO 6 — EXCEL
Abas: {', '.join(abas_excel)}
Colunas extras: {cols_extras}
"""
        st.session_state.briefing_atual = briefing

        system_prompt = """Você é um especialista em escalas de internato médico.
Siga EXATAMENTE o PROMPT MESTRE abaixo:

FASE 1: Confirme que entendeu o briefing listando alunos, rodízio e regras.
FASE 2: Construa a alocação passo a passo (manhãs → tardes → terças → FDS → compensações → ajuste CH).
FASE 3: Auditoria — verifique CH, cobertura mínima, regras FDS.
FASE 4: Gere o resultado em JSON.

REGRAS OBRIGATÓRIAS:
- Nunca ultrapasse o limite absoluto de CH
- Respeite cobertura mínima por turno
- Quinta sem tarde (ENAMED) salvo indicação contrária
- Terça com tarde encurtada salvo indicação contrária
- FDS: quem faz e compensação conforme briefing
- Distribua FDS homogeneamente (range máx 1 entre alunos do bloco)

Responda em JSON com esta estrutura exata:
{
  "confirmacao": "resumo do que entendeu",
  "calendario_rodizio": [
    {"semana": 1, "periodo": "DD/MM-DD/MM", "alocacao": {"SG1": "Local", "SG2": "Local"}}
  ],
  "escala_detalhada": [
    {"semana": 1, "data": "DD/MM", "dia": "Seg", "local": "Enf", "turno": "Manhã", "horario": "07-13h", "horas": 6, "alunos": ["Nome1"]}
  ],
  "resumo_horas": [
    {"sg": 1, "nome": "Nome", "total_horas": 40, "semanas": [40,40,40,40,40,40,40,40], "plantoes_fds": 1}
  ],
  "auditoria": {
    "ch_ok": true, "cobertura_ok": true, "fds_ok": true,
    "erros": [], "avisos": [], "aprovado": true
  }
}"""

        with st.spinner("IA gerando sua escala completa... ⏳ Pode levar até 2 minutos"):
            resposta = chamar_claude(
                [{"role": "user", "content": f"Gere a escala completa:\n\n{briefing}"}],
                system_prompt=system_prompt, max_tokens=8000
            )
            if resposta:
                st.session_state.escala_gerada = resposta
                st.session_state.esp_atual = especialidade
                st.session_state.grupo_atual = grupo
                st.session_state.turma_atual = turma
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

def _mostrar_resultado(resposta_raw, esp, grupo, turma):
    """Mostra o resultado da IA e oferece download."""
    try:
        txt = resposta_raw.strip()
        if "```" in txt:
            partes = txt.split("```")
            for p in partes:
                if p.startswith("json"): txt = p[4:]; break
                elif "{" in p: txt = p; break

        dados = json.loads(txt)

        # Confirmação
        if dados.get("confirmacao"):
            with st.expander("📋 O que a IA entendeu", expanded=False):
                st.write(dados["confirmacao"])

        # Auditoria
        audit = dados.get("auditoria", {})
        if audit.get("aprovado"):
            st.success("✅ Auditoria aprovada! Escala sem erros.")
        else:
            for err in audit.get("erros", []):
                st.error(f"❌ {err}")
        for av in audit.get("avisos", []):
            st.warning(f"⚠️ {av}")

        # Calendário
        if dados.get("calendario_rodizio"):
            st.subheader("📅 Calendário de Rodízio")
            cal_data = dados["calendario_rodizio"]
            df_cal = pd.DataFrame([
                {"Semana": f"Sem {r['semana']} | {r['periodo']}",
                 **{k: v for k,v in r.get("alocacao",{}).items()}}
                for r in cal_data
            ])
            st.dataframe(df_cal, use_container_width=True, hide_index=True)

        # Resumo de horas
        if dados.get("resumo_horas"):
            st.subheader("⏱️ Resumo de Horas por Aluno")
            df_h = pd.DataFrame(dados["resumo_horas"])
            if "sg" in df_h.columns: df_h = df_h.sort_values(["sg","nome"])
            st.dataframe(df_h, use_container_width=True, hide_index=True)

            # Verificar CH fora do limite
            if "total_horas" in df_h.columns:
                fora = df_h[df_h["total_horas"] > 43]
                if not fora.empty:
                    st.warning(f"⚠️ {len(fora)} aluno(s) com CH acima de 43h!")

        # Downloads
        st.subheader("📥 Exportar Escala")
        col_dl1, col_dl2, col_dl3 = st.columns(3)

        # Excel — 9 abas formato CM5
        try:
            from gerador_excel_final import gerar_excel_completo
            config_excel = st.session_state.get("config_atual", {})
            config_excel.update({"especialidade": esp, "grupo": grupo, "turma": turma})
            xlsx_bytes = gerar_excel_completo(dados, config_excel)
            with col_dl1:
                st.download_button(
                    "📊 Excel completo (9 abas)",
                    data=xlsx_bytes,
                    file_name=f"Escala_{esp}_{grupo}_{turma}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        except Exception as e_excel:
            with col_dl1:
                st.warning(f"Erro Excel: {e_excel}")

        # CSV — Escala nominal detalhada
        if dados.get("escala_detalhada"):
            import csv, io as io_csv
            ed = dados["escala_detalhada"]
            buf = io_csv.StringIO()
            fieldnames = ["semana","data","dia","local","turno","horario","horas","sg","nome","ra"]
            writer = csv.DictWriter(buf, fieldnames=fieldnames, extrasaction="ignore")
            writer.writeheader()
            for entry in ed:
                alunos = entry.get("alunos", [])
                nome = entry.get("nome","") if isinstance(entry.get("nome"), str) else ""
                if alunos and not nome:
                    for a in alunos:
                        row_dict = {k: entry.get(k,"") for k in fieldnames}
                        row_dict["nome"] = a
                        writer.writerow(row_dict)
                else:
                    writer.writerow({k: entry.get(k,"") for k in fieldnames})
            with col_dl2:
                st.download_button(
                    "📄 CSV — Escala detalhada",
                    data=buf.getvalue().encode("utf-8-sig"),
                    file_name=f"Escala_{esp}_{grupo}_{turma}_nominal.csv",
                    mime="text/csv",
                    use_container_width=True
                )

        # CSV — Resumo de horas
        if dados.get("resumo_horas"):
            buf2 = io_csv.StringIO()
            rh = dados["resumo_horas"]
            if rh:
                fields2 = list(rh[0].keys())
                writer2 = csv.DictWriter(buf2, fieldnames=fields2, extrasaction="ignore")
                writer2.writeheader()
                for r in rh:
                    row2 = {}
                    for k in fields2:
                        v = r.get(k,"")
                        row2[k] = " | ".join(str(x) for x in v) if isinstance(v, list) else v
                    writer2.writerow(row2)
            with col_dl3:
                st.download_button(
                    "📄 CSV — Resumo de horas",
                    data=buf2.getvalue().encode("utf-8-sig"),
                    file_name=f"Escala_{esp}_{grupo}_{turma}_horas.csv",
                    mime="text/csv",
                    use_container_width=True
                )

    except json.JSONDecodeError:
        st.subheader("📄 Resposta da IA (texto)")
        st.text_area("Resultado", resposta_raw, height=400)
        st.warning("A IA respondeu em texto — o Excel não pôde ser gerado automaticamente. Copie o conteúdo acima.")

# Mostrar resultado se já gerado
if "escala_gerada" in st.session_state:
    st.divider()
    st.header("📊 Resultado")
    _mostrar_resultado(
        st.session_state.escala_gerada,
        st.session_state.get("esp_atual", especialidade if 'especialidade' in dir() else ""),
        st.session_state.get("grupo_atual", ""),
        st.session_state.get("turma_atual", "")
    )

    if st.button("🔁 Solicitar correção à IA"):
        correcao = st.text_area("Descreva o que precisa corrigir:", key="correcao_txt")
        if correcao and st.button("Enviar correção", key="btn_correcao"):
            with st.spinner("IA corrigindo..."):
                resposta2 = chamar_claude([
                    {"role": "user", "content": f"Gere a escala:\n{st.session_state.get('briefing_atual','')}"},
                    {"role": "assistant", "content": st.session_state.escala_gerada},
                    {"role": "user", "content": f"Corrija o seguinte:\n{correcao}\n\nRetorne o JSON completo corrigido."}
                ], max_tokens=8000)
                if resposta2:
                    st.session_state.escala_gerada = resposta2
                    st.rerun()
