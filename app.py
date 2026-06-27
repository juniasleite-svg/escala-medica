import streamlit as st
import pandas as pd
from gerador_escala import gerar_escala
import io
import json
import datetime

st.set_page_config(page_title="Gerador de Escalas Médicas", page_icon="🏥", layout="wide")
st.title("🏥 Gerador de Escalas Médicas")

# ── MODO DE USO ──────────────────────────────────────────────────────────────
modo = st.radio(
    "Como deseja começar?",
    ["✨ Criar nova escala", "📂 Importar escala existente e modificar"],
    horizontal=True
)

st.divider()

# ════════════════════════════════════════════════════════════════════════════
# MODO: IMPORTAR ESCALA EXISTENTE
# ════════════════════════════════════════════════════════════════════════════
if modo == "📂 Importar escala existente e modificar":
    st.header("📂 Importar Escala Existente")
    st.caption("Faça upload de qualquer Excel de escala já feita — a IA vai ler e entender a lógica automaticamente.")

    col_up1, col_up2 = st.columns(2)
    with col_up1:
        arquivo_escala = st.file_uploader("Upload da escala existente (.xlsx)", type=["xlsx"], key="import_escala")
    with col_up2:
        arquivo_alunos_imp = st.file_uploader("Upload da base de alunos (opcional)", type=["xlsx"], key="import_alunos")

    if arquivo_escala:
        if st.button("🤖 Analisar escala com IA", type="primary", use_container_width=True):
            with st.spinner("A IA está lendo e interpretando sua escala... ⏳"):
                try:
                    # Ler todas as abas do Excel
                    xls = pd.ExcelFile(arquivo_escala)
                    conteudo_excel = ""
                    for sheet in xls.sheet_names:
                        df = pd.read_excel(arquivo_escala, sheet_name=sheet, header=None)
                        conteudo_excel += f"\n\n=== ABA: {sheet} ===\n"
                        conteudo_excel += df.fillna("").to_string(index=False, header=False)
                        if len(conteudo_excel) > 8000:
                            conteudo_excel = conteudo_excel[:8000] + "\n...[truncado]"
                            break

                    prompt = f"""Você é um especialista em escalas de internato médico.
Analise o conteúdo deste Excel de escala médica e extraia as informações em formato JSON.

CONTEÚDO DO EXCEL:
{conteudo_excel}

Retorne APENAS um JSON válido com esta estrutura (sem explicações, sem markdown):
{{
  "especialidade": "nome da especialidade",
  "grupo": "ex: GRUPO D",
  "turma": "ex: T6",
  "ano_curso": "ex: 4º Ano",
  "data_inicio": "YYYY-MM-DD",
  "num_semanas": 8,
  "num_sg": 6,
  "locais": [
    {{
      "nome": "Nome do local",
      "tipo": "padrao ou bloco_ped",
      "fds": false,
      "turno_m": true,
      "turno_t": true,
      "turno_c": false,
      "vagas_m": 6,
      "vagas_t": 6,
      "vagas_c": 0
    }}
  ],
  "rodizio": {{
    "descricao": "ex: rodízio 3-3-2, cada SG fica X semanas em cada local",
    "pares_bp": {{}},
    "rotacao_bp": "3-3-2"
  }},
  "bloqueios": {{
    "quinta_tarde": true,
    "terca_parcial": false,
    "observacoes": "outras regras identificadas"
  }},
  "subgrupos": [
    {{"sg": 1, "alunos": ["Nome 1", "Nome 2"]}}
  ],
  "resumo": "Descrição em português do que você entendeu sobre a lógica desta escala"
}}"""

                    import requests
                    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
                    response = requests.post(
                        "https://api.anthropic.com/v1/messages",
                        headers={
                            "x-api-key": api_key,
                            "anthropic-version": "2023-06-01",
                            "content-type": "application/json"
                        },
                        json={
                            "model": "claude-sonnet-4-6",
                            "max_tokens": 4000,
                            "messages": [{"role": "user", "content": prompt}]
                        },
                        timeout=60
                    )

                    result = response.json()
                    texto = result["content"][0]["text"].strip()
                    # Limpar markdown se vier
                    if "```" in texto:
                        texto = texto.split("```")[1]
                        if texto.startswith("json"):
                            texto = texto[4:]
                    
                    dados = json.loads(texto)
                    st.session_state.escala_importada = dados
                    st.success("✅ Escala analisada com sucesso!")

                except Exception as e:
                    st.error(f"Erro ao analisar: {e}")

        if "escala_importada" in st.session_state:
            dados = st.session_state.escala_importada

            st.success("✅ Escala analisada! Confira abaixo e edite o que precisar.")
            st.info(f"💡 **Resumo da IA:** {dados.get('resumo','')}")
            st.divider()

            # ── Identificação (pré-preenchida, editável) ─────────────────────
            st.subheader("1️⃣ Identificação da Escala")
            col_i1, col_i2 = st.columns(2)
            with col_i1:
                esp_imp = st.text_input("Especialidade", value=dados.get("especialidade",""), key="imp_esp")
                ano_imp = st.selectbox("Ano do curso", ["3º Ano","4º Ano","5º Ano","6º Ano"],
                                        index=["3º Ano","4º Ano","5º Ano","6º Ano"].index(dados.get("ano_curso","4º Ano")) 
                                        if dados.get("ano_curso") in ["3º Ano","4º Ano","5º Ano","6º Ano"] else 1,
                                        key="imp_ano")
            with col_i2:
                grupo_imp = st.text_input("Grupo", value=dados.get("grupo",""), key="imp_grupo")
                turma_imp = st.text_input("Turma", value=dados.get("turma",""), key="imp_turma")
            col_i3, col_i4 = st.columns(2)
            with col_i3:
                try:
                    data_imp_default = datetime.date.fromisoformat(dados.get("data_inicio", str(datetime.date.today())))
                except:
                    data_imp_default = datetime.date.today()
                data_imp = st.date_input("Data de início", value=data_imp_default, key="imp_data_inicio")
            with col_i4:
                nsem_imp = st.number_input("Número de semanas", min_value=1, max_value=20,
                                            value=int(dados.get("num_semanas", 8)), key="imp_nsem2")

            # ── Locais (pré-preenchidos, editáveis) ──────────────────────────
            st.subheader("2️⃣ Locais de Rodízio")
            locais_imp = dados.get("locais", [])
            locais_editados = []
            for i, local in enumerate(locais_imp):
                with st.expander(f"📍 Local {i+1}: {local.get('nome','')}", expanded=False):
                    col_l1, col_l2 = st.columns(2)
                    with col_l1:
                        nome_l = st.text_input("Nome", value=local.get("nome",""), key=f"imp_lnome_{i}")
                        fds_l = st.checkbox("Fim de semana", value=local.get("fds", False), key=f"imp_lfds_{i}")
                    with col_l2:
                        tm_l = st.checkbox("Manhã (7-13h)", value=local.get("turno_m", True), key=f"imp_lm_{i}")
                        tt_l = st.checkbox("Tarde (13-19h)", value=local.get("turno_t", True), key=f"imp_lt_{i}")
                        tc_l = st.checkbox("Cinderela (19-23h)", value=local.get("turno_c", False), key=f"imp_lc_{i}")
                    col_v1, col_v2, col_v3 = st.columns(3)
                    with col_v1: vm_l = st.number_input("Vagas manhã", 0, 20, int(local.get("vagas_m",6)), key=f"imp_lvm_{i}") if tm_l else 0
                    with col_v2: vt_l = st.number_input("Vagas tarde", 0, 20, int(local.get("vagas_t",6)), key=f"imp_lvt_{i}") if tt_l else 0
                    with col_v3: vc_l = st.number_input("Vagas cinderela", 0, 10, int(local.get("vagas_c",2)), key=f"imp_lvc_{i}") if tc_l else 0
                    locais_editados.append({"nome": nome_l, "tipo": local.get("tipo","padrao"), "fds": fds_l,
                                             "turno_m": tm_l, "vagas_m": int(vm_l),
                                             "turno_t": tt_l, "vagas_t": int(vt_l),
                                             "turno_c": tc_l, "vagas_c": int(vc_l)})

            # ── Alunos (pré-preenchidos, editáveis) ──────────────────────────
            st.subheader("3️⃣ Subgrupos e Alunos")
            subgrupos_imp = dados.get("subgrupos", [])
            alunos_editados = []
            for sg_info in subgrupos_imp:
                sg_num = sg_info.get("sg", 1)
                with st.expander(f"👥 SG{sg_num} ({len(sg_info.get('alunos',[]))} alunos)", expanded=False):
                    alunos_sg_txt = st.text_area(f"Alunos do SG{sg_num} (um por linha)",
                                                  value="\n".join(sg_info.get("alunos",[])),
                                                  key=f"imp_sg_{sg_num}", height=120)
                    for nome_a in alunos_sg_txt.strip().split("\n"):
                        if nome_a.strip():
                            alunos_editados.append({"Nome Completo": nome_a.strip(), "RA": "", "Sub Grupo": sg_num})

            # ── Bloqueios (pré-preenchidos, editáveis) ────────────────────────
            st.subheader("4️⃣ Bloqueios")
            bloq_imp = dados.get("bloqueios", {})
            col_bq1, col_bq2 = st.columns(2)
            with col_bq1:
                quinta_imp = st.checkbox("Quinta-feira tarde bloqueada", value=bloq_imp.get("quinta_tarde", True), key="imp_quinta2")
            with col_bq2:
                terca_imp = st.checkbox("Terça-feira: campo só até 16h", value=bloq_imp.get("terca_parcial", False), key="imp_terca2")
            if bloq_imp.get("observacoes"):
                st.caption(f"ℹ️ Outras regras identificadas: {bloq_imp.get('observacoes','')}")

            st.divider()
            modificacoes = []  # sem multiselect — edição direta nos campos acima

            if "ajustes_import" not in st.session_state:
                st.session_state.ajustes_import = []

            # ── Trocar aluno / Substituir ausente ───────────────────────────
            if "Trocar aluno de subgrupo" in modificacoes or "Substituir aluno ausente" in modificacoes:
                with st.expander("👤 Ajustes de alunos", expanded=True):
                    todos_alunos = []
                    for sg_info in dados.get("subgrupos",[]):
                        todos_alunos.extend(sg_info.get("alunos",[]))
                    col_m1, col_m2, col_m3 = st.columns(3)
                    with col_m1:
                        aluno_sel = st.selectbox("Aluno", sorted(todos_alunos) if todos_alunos else ["(nenhum identificado)"], key="imp_aluno")
                    with col_m2:
                        tipo_mod = st.selectbox("Tipo", ["Trocar de SG", "Marcar ausente na semana"], key="imp_tipo")
                    with col_m3:
                        sem_mod = st.number_input("Semana", min_value=1, max_value=int(dados.get("num_semanas",8)), value=1, key="imp_sem")
                    if tipo_mod == "Trocar de SG":
                        novo_sg_mod = st.number_input("Novo SG", min_value=1, max_value=int(dados.get("num_sg",6)), value=1, key="imp_sg")
                    if st.button("➕ Adicionar ajuste de aluno", key="btn_aj_aluno"):
                        aj = {"tipo_ajuste": "aluno", "aluno": aluno_sel, "acao": tipo_mod, "semana": int(sem_mod)}
                        if tipo_mod == "Trocar de SG": aj["novo_sg"] = int(novo_sg_mod)
                        st.session_state.ajustes_import.append(aj)
                        st.success("Ajuste adicionado!")
                    for aj in [a for a in st.session_state.ajustes_import if a.get("tipo_ajuste")=="aluno"]:
                        st.write(f"• **{aj['aluno']}** | {aj['acao']} | Sem {aj['semana']}")

            # ── Mudar turnos de um local ─────────────────────────────────────
            if "Mudar turnos de um local" in modificacoes:
                with st.expander("🕐 Mudar turnos de um local", expanded=True):
                    locais_import = [l["nome"] for l in dados.get("locais",[])]
                    if not locais_import:
                        locais_import = ["(locais não identificados — digite o nome)"]
                    col_t1, col_t2 = st.columns(2)
                    with col_t1:
                        local_turno = st.selectbox("Local", locais_import, key="imp_local_turno")
                        sem_turno = st.number_input("Aplicar a partir da semana", min_value=1, 
                                                     max_value=int(dados.get("num_semanas",8)), value=1, key="imp_sem_turno")
                    with col_t2:
                        novo_m = st.checkbox("Manhã (7-13h)", value=True, key="imp_tm")
                        novo_t = st.checkbox("Tarde (13-19h)", value=True, key="imp_tt")
                        novo_c = st.checkbox("Cinderela (19-23h)", key="imp_tc")
                        novo_fds = st.checkbox("Fim de semana", key="imp_fds")
                    col_v1, col_v2, col_v3 = st.columns(3)
                    with col_v1: vagas_m_novo = st.number_input("Vagas manhã", 0, 20, 6, key="imp_vm") if novo_m else 0
                    with col_v2: vagas_t_novo = st.number_input("Vagas tarde", 0, 20, 6, key="imp_vt") if novo_t else 0
                    with col_v3: vagas_c_novo = st.number_input("Vagas cinderela", 0, 10, 2, key="imp_vc") if novo_c else 0
                    if st.button("➕ Confirmar mudança de turno", key="btn_turno"):
                        aj = {"tipo_ajuste": "turno", "local": local_turno, "a_partir_semana": int(sem_turno),
                              "turno_m": novo_m, "vagas_m": int(vagas_m_novo),
                              "turno_t": novo_t, "vagas_t": int(vagas_t_novo),
                              "turno_c": novo_c, "vagas_c": int(vagas_c_novo), "fds": novo_fds}
                        st.session_state.ajustes_import.append(aj)
                        st.success(f"Mudança de turno adicionada para {local_turno} a partir da semana {sem_turno}!")
                    for aj in [a for a in st.session_state.ajustes_import if a.get("tipo_ajuste")=="turno"]:
                        turnos = []
                        if aj["turno_m"]: turnos.append("Manhã")
                        if aj["turno_t"]: turnos.append("Tarde")
                        if aj["turno_c"]: turnos.append("Cinderela")
                        st.write(f"• **{aj['local']}** | {' + '.join(turnos)} | A partir sem {aj['a_partir_semana']}")

            # ── Adicionar local ──────────────────────────────────────────────
            if "Adicionar local" in modificacoes:
                with st.expander("📍 Adicionar local", expanded=True):
                    col_al1, col_al2 = st.columns(2)
                    with col_al1:
                        nome_novo_local = st.text_input("Nome do novo local", key="imp_nome_local")
                        sem_inicio_local = st.number_input("A partir da semana", min_value=1, max_value=int(dados.get("num_semanas",8)), value=1, key="imp_sem_local")
                    with col_al2:
                        nl_m = st.checkbox("Manhã", value=True, key="nl_m")
                        nl_t = st.checkbox("Tarde", value=True, key="nl_t")
                        nl_c = st.checkbox("Cinderela", key="nl_c")
                        nl_fds = st.checkbox("Fim de semana", key="nl_fds")
                    if st.button("➕ Adicionar este local", key="btn_add_local"):
                        aj = {"tipo_ajuste": "add_local", "nome": nome_novo_local, "a_partir": int(sem_inicio_local),
                              "turno_m": nl_m, "turno_t": nl_t, "turno_c": nl_c, "fds": nl_fds}
                        st.session_state.ajustes_import.append(aj)
                        st.success(f"Local '{nome_novo_local}' adicionado!")

            # ── Remover local ────────────────────────────────────────────────
            if "Remover local" in modificacoes:
                with st.expander("🗑️ Remover local", expanded=True):
                    locais_import2 = [l["nome"] for l in dados.get("locais",[])]
                    local_remover = st.selectbox("Local a remover", locais_import2 if locais_import2 else ["(nenhum)"], key="imp_rem_local")
                    sem_rem = st.number_input("A partir da semana", min_value=1, max_value=int(dados.get("num_semanas",8)), value=1, key="imp_rem_sem")
                    if st.button("🗑️ Confirmar remoção", key="btn_rem_local"):
                        aj = {"tipo_ajuste": "rem_local", "nome": local_remover, "a_partir": int(sem_rem)}
                        st.session_state.ajustes_import.append(aj)
                        st.success(f"Remoção de '{local_remover}' agendada a partir da semana {sem_rem}!")

            # ── Alterar datas ────────────────────────────────────────────────
            if "Alterar datas" in modificacoes:
                with st.expander("📅 Alterar datas", expanded=True):
                    nova_data = st.date_input("Nova data de início", key="imp_data")
                    if st.button("✅ Confirmar nova data", key="btn_data"):
                        st.session_state.ajustes_import.append({"tipo_ajuste": "data", "nova_data": str(nova_data)})
                        st.success(f"Nova data: {nova_data.strftime('%d/%m/%Y')}")

            # ── Mudar número de semanas ──────────────────────────────────────
            if "Mudar número de semanas" in modificacoes:
                with st.expander("🗓️ Mudar número de semanas", expanded=True):
                    novo_num_sem = st.number_input("Novo número de semanas", min_value=1, max_value=20, 
                                                    value=int(dados.get("num_semanas",8)), key="imp_nsem")
                    if st.button("✅ Confirmar", key="btn_nsem"):
                        st.session_state.ajustes_import.append({"tipo_ajuste": "num_semanas", "valor": int(novo_num_sem)})
                        st.success(f"Número de semanas alterado para {novo_num_sem}!")

            # ── Alterar bloqueios ────────────────────────────────────────────
            if "Alterar bloqueios fixos" in modificacoes:
                with st.expander("🚫 Alterar bloqueios fixos", expanded=True):
                    novo_quinta = st.checkbox("Quinta-feira tarde bloqueada", value=dados.get("bloqueios",{}).get("quinta_tarde",True), key="imp_quinta")
                    novo_terca = st.checkbox("Terça-feira: campo só até 16h", value=dados.get("bloqueios",{}).get("terca_parcial",False), key="imp_terca")
                    novo_feriado = st.text_area("Feriados extras (DD/MM/AAAA, um por linha)", key="imp_feriados")
                    if st.button("✅ Confirmar bloqueios", key="btn_bloq"):
                        st.session_state.ajustes_import.append({
                            "tipo_ajuste": "bloqueios", "quinta_tarde": novo_quinta,
                            "terca_parcial": novo_terca, "feriados_extras": novo_feriado
                        })
                        st.success("Bloqueios atualizados!")

            # ── Resumo dos ajustes ───────────────────────────────────────────
            if st.session_state.ajustes_import:
                st.subheader(f"📝 {len(st.session_state.ajustes_import)} ajuste(s) cadastrado(s)")
                for i, aj in enumerate(st.session_state.ajustes_import):
                    col_x, col_desc = st.columns([1,10])
                    with col_x:
                        if st.button("❌", key=f"del_imp_{i}"):
                            st.session_state.ajustes_import.pop(i); st.rerun()
                    with col_desc:
                        st.write(f"• {aj}")

            if st.button("🚀 Gerar Escala com Modificações", type="primary", use_container_width=True):
                if not alunos_editados:
                    st.error("Nenhum aluno identificado!")
                elif not locais_editados:
                    st.error("Nenhum local identificado!")
                else:
                    num_sg_imp = max([a["Sub Grupo"] for a in alunos_editados]) if alunos_editados else 1
                    config_imp = {
                        "especialidade": esp_imp, "ano_curso": ano_imp,
                        "grupo": grupo_imp, "turma": turma_imp,
                        "data_inicio": str(data_imp), "num_semanas": int(nsem_imp),
                        "num_sg": num_sg_imp, "locais": locais_editados,
                        "pares_bp": {}, "rotacao_bp": "3-3-2",
                        "bloqueios": {"quinta_tarde": quinta_imp, "terca_parcial": terca_imp, "feriados": []},
                        "ajustes": [], "alunos": alunos_editados,
                    }
                    with st.spinner("Gerando escala... ⏳"):
                        try:
                            from gerador_escala import gerar_escala
                            resultado = gerar_escala(config_imp)
                            st.success("✅ Escala gerada!")
                            col_dl1, col_dl2 = st.columns(2)
                            with col_dl1:
                                st.download_button("📥 Baixar Excel (.xlsx)", data=resultado["xlsx"],
                                    file_name=f"Escala_{esp_imp}_{grupo_imp}_{turma_imp}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True)
                            with col_dl2:
                                st.download_button("📥 Baixar PowerPoint (.pptx)", data=resultado["pptx"],
                                    file_name=f"Escala_{esp_imp}_{grupo_imp}_{turma_imp}.pptx",
                                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                    use_container_width=True)
                        except Exception as e:
                            st.error(f"Erro: {e}"); st.exception(e)

    st.stop()

# ════════════════════════════════════════════════════════════════════════════
# MODO: CRIAR NOVA ESCALA
# ════════════════════════════════════════════════════════════════════════════

# ── PASSO 0: Upload do Excel ─────────────────────────────────────────────────
st.header("Passo 0 — Base de Alunos")
arquivo_alunos = st.file_uploader(
    "Faça upload do Excel com grupos e subgrupos (TODOS_GRUPOS_SUBGRUPOS...xlsx)",
    type=["xlsx"],
    help="Deve ter abas por grupo (GRUPO A, GRUPO B...) com colunas: Nome Completo, RA, Grupo, OPÇÃO, Sub Grupo"
)

grupos_disponiveis = []
if arquivo_alunos:
    try:
        xls = pd.ExcelFile(arquivo_alunos)
        grupos_disponiveis = [s for s in xls.sheet_names if s.upper().startswith("GRUPO")]
        st.success(f"✅ Arquivo carregado! Grupos encontrados: {', '.join(grupos_disponiveis)}")
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")

# ── PASSO 1: Identificação ───────────────────────────────────────────────────
st.header("Passo 1 — Identificação da Escala")

col1, col2 = st.columns(2)
with col1:
    especialidade = st.text_input("Especialidade *", placeholder="ex: Clínica Médica, Pediatria, Cirurgia...")
    ano_curso = st.selectbox("Ano do curso *", ["3º Ano", "4º Ano", "5º Ano", "6º Ano"], index=1)
with col2:
    grupo_opcoes = grupos_disponiveis if grupos_disponiveis else ["GRUPO A","GRUPO B","GRUPO C","GRUPO D","GRUPO E","GRUPO F"]
    grupo_selecionado = st.selectbox("Grupo *", grupo_opcoes)
    turma = st.text_input("Turma *", placeholder="ex: T6")

col3, col4 = st.columns(2)
with col3:
    data_inicio = st.date_input("Data de início (segunda-feira) *")
with col4:
    num_semanas = st.number_input("Número de semanas *", min_value=1, max_value=20, value=8)

# ── PASSO 2: Locais de Rodízio ───────────────────────────────────────────────
st.header("Passo 2 — Locais de Rodízio")

num_locais = st.number_input("Quantos locais de estágio?", min_value=2, max_value=6, value=3)

def calcular_sg_ideal(n_locais):
    return {2:4, 3:6, 4:4, 5:5, 6:6}.get(int(n_locais), 4)

sg_sugerido = calcular_sg_ideal(num_locais)
st.info(f"💡 Com {int(num_locais)} locais, o sistema sugere **{sg_sugerido} subgrupos**.")

is_pediatria = "pediatria" in especialidade.lower() if especialidade else False
locais = []
tem_bloco_ped = False

for i in range(int(num_locais)):
    with st.expander(f"Local {i+1}", expanded=(i==0)):
        col_a, col_b = st.columns(2)
        with col_a:
            nome_local = st.text_input(f"Nome do local {i+1}", key=f"local_nome_{i}", placeholder="ex: UPA Mandic")
        with col_b:
            tipo_local = st.selectbox(
                f"Tipo do local {i+1}",
                ["Padrão", "Bloco Pediátrico (Enfermaria + PA)"] if is_pediatria else ["Padrão"],
                key=f"tipo_{i}"
            )
        eh_bloco_ped = tipo_local == "Bloco Pediátrico (Enfermaria + PA)"
        if eh_bloco_ped:
            tem_bloco_ped = True
        tem_fds = st.checkbox("Funciona no fim de semana?", key=f"fds_{i}", value=eh_bloco_ped)
        if eh_bloco_ped:
            st.markdown("**🏥 Configuração do Bloco Pediátrico**")
            col_pa1, col_pa2 = st.columns(2)
            with col_pa1:
                vagas_enf = st.number_input("Vagas Enfermaria (manhã)", min_value=1, max_value=10, value=4, key=f"enf_{i}")
            with col_pa2:
                vagas_pa_turno = st.number_input("Vagas PA por turno (máx)", min_value=1, max_value=4, value=2, key=f"pa_{i}")
            locais.append({"nome": nome_local, "tipo": "bloco_ped", "fds": True,
                           "vagas_enf": int(vagas_enf), "vagas_pa": int(vagas_pa_turno),
                           "turno_m": True, "turno_t": True, "turno_c": True,
                           "vagas_m": int(vagas_enf), "vagas_t": 0, "vagas_c": 0})
        else:
            col_c, col_d, col_e = st.columns(3)
            with col_c: turno_m = st.checkbox("Manhã (7-13h)", value=True, key=f"m_{i}")
            with col_d: turno_t = st.checkbox("Tarde (13-19h)", value=True, key=f"t_{i}")
            with col_e: turno_c = st.checkbox("Cinderela (19-23h)", key=f"c_{i}")
            col_f, col_g, col_h = st.columns(3)
            with col_f: vagas_m = st.number_input("Vagas manhã", min_value=0, max_value=20, value=6, key=f"vm_{i}") if turno_m else 0
            with col_g: vagas_t = st.number_input("Vagas tarde", min_value=0, max_value=20, value=6, key=f"vt_{i}") if turno_t else 0
            with col_h: vagas_c = st.number_input("Vagas cinderela", min_value=0, max_value=10, value=2, key=f"vc_{i}") if turno_c else 0
            locais.append({"nome": nome_local, "tipo": "padrao", "fds": tem_fds,
                           "turno_m": turno_m, "vagas_m": int(vagas_m),
                           "turno_t": turno_t, "vagas_t": int(vagas_t),
                           "turno_c": turno_c, "vagas_c": int(vagas_c)})

# ── PASSO 2B: Bloco Pediátrico ───────────────────────────────────────────────
pares_bp = {}
rotacao_bp = "3-3-2"

if tem_bloco_ped:
    st.header("Passo 2B — Configuração do Bloco Pediátrico")
    col_bp1, col_bp2 = st.columns(2)
    with col_bp1:
        rotacao_bp = st.selectbox("Rodízio no Bloco Pediátrico", ["3-3-2","4-4","2-2-2-2","Personalizado"])
    with col_bp2:
        num_pares = st.number_input("Pares de SGs simultâneos no BP", min_value=1, max_value=3, value=2)
    
    if rotacao_bp == "3-3-2": semanas_por_par = [3,3,2]
    elif rotacao_bp == "4-4": semanas_por_par = [4,4]
    elif rotacao_bp == "2-2-2-2": semanas_por_par = [2,2,2,2]
    else: semanas_por_par = []

    meio = sg_sugerido // 2
    pares_sugeridos = [(p+1, p+meio+1) for p in range(int(num_pares)) if p+meio+1 <= sg_sugerido]
    
    sem_acumulada = 0
    for idx, sem_qtd in enumerate(semanas_por_par):
        col_p1, col_p2, col_p3 = st.columns([2,2,3])
        with col_p1:
            sg_a_val = st.number_input(f"Par {idx+1} — SG A", min_value=1, max_value=sg_sugerido,
                                        value=pares_sugeridos[idx%len(pares_sugeridos)][0] if pares_sugeridos else 1,
                                        key=f"par_a_{idx}")
        with col_p2:
            sg_b_val = st.number_input(f"Par {idx+1} — SG B", min_value=1, max_value=sg_sugerido,
                                        value=pares_sugeridos[idx%len(pares_sugeridos)][1] if pares_sugeridos else 2,
                                        key=f"par_b_{idx}")
        with col_p3:
            st.info(f"Semanas {sem_acumulada+1} a {sem_acumulada+sem_qtd}")
        for s in range(sem_acumulada, sem_acumulada+sem_qtd):
            pares_bp[s] = (int(sg_a_val), int(sg_b_val))
        sem_acumulada += sem_qtd

# ── PASSO 3: Bloqueios ───────────────────────────────────────────────────────
st.header("Passo 3 — Bloqueios e Feriados")

col_b1, col_b2 = st.columns(2)
with col_b1: bloquear_quinta_tarde = st.checkbox("Quinta-feira tarde bloqueada (ENAMED)", value=True)
with col_b2: bloquear_terca_parcial = st.checkbox("Terça-feira: campo só até 16h", value=True)

feriados_nacionais_2026 = [
    datetime.date(2026,7,9), datetime.date(2026,9,7), datetime.date(2026,10,12),
    datetime.date(2026,11,2), datetime.date(2026,11,15), datetime.date(2026,11,20),
    datetime.date(2026,12,25),
]
col_f1, col_f2 = st.columns(2)
with col_f1: usar_feriados_auto = st.checkbox("Feriados nacionais + SP automáticos", value=True)
with col_f2:
    if usar_feriados_auto:
        st.caption("Incluídos: " + ", ".join([f.strftime("%d/%m") for f in feriados_nacionais_2026]))

feriados_extras_str = st.text_area("Feriados extras (DD/MM/AAAA, um por linha)", placeholder="ex:\n09/07/2026")
feriados_finais = list(feriados_nacionais_2026) if usar_feriados_auto else []
if feriados_extras_str:
    for linha in feriados_extras_str.strip().split("\n"):
        try: feriados_finais.append(datetime.datetime.strptime(linha.strip(), "%d/%m/%Y").date())
        except: pass

# ── PASSO 4: Alunos ──────────────────────────────────────────────────────────
st.header("Passo 4 — Alunos")
alunos_finais = []

if arquivo_alunos and grupo_selecionado:
    try:
        df_grupo = pd.read_excel(arquivo_alunos, sheet_name=grupo_selecionado)
        opcao_col = "OPÇÃO"
        opcoes_disponiveis = df_grupo[opcao_col].dropna().unique() if opcao_col in df_grupo.columns else []
        opcao_escolhida = next((op for op in opcoes_disponiveis if f"{sg_sugerido} SG" in str(op)), 
                                opcoes_disponiveis[0] if len(opcoes_disponiveis) > 0 else None)
        if opcao_escolhida is not None:
            opcao_selecionada = st.selectbox("Opção de subgrupos", opcoes_disponiveis,
                                              index=list(opcoes_disponiveis).index(opcao_escolhida))
            df_filtrado = df_grupo[df_grupo[opcao_col] == opcao_selecionada].copy()
            st.success(f"✅ **{len(df_filtrado)} alunos** — {opcao_selecionada}")
            df_display = df_filtrado[["Nome Completo","RA","Sub Grupo"]].copy()
            df_display["Sub Grupo"] = df_display["Sub Grupo"].astype(int)
            st.dataframe(df_display.sort_values("Sub Grupo"), use_container_width=True, hide_index=True)
            alunos_finais = df_filtrado[["Nome Completo","RA","Sub Grupo"]].to_dict("records")
    except Exception as e:
        st.error(f"Erro: {e}")
else:
    st.info("👆 Faça o upload do Excel no Passo 0 para carregar os alunos automaticamente.")

# ── PASSO 5: Ajustes ────────────────────────────────────────────────────────
if alunos_finais:
    st.header("Passo 5 — Ajustes Individuais (Opcional)")
    with st.expander("➕ Adicionar ajuste"):
        nomes = [a["Nome Completo"] for a in alunos_finais]
        col_aj1, col_aj2, col_aj3 = st.columns(3)
        with col_aj1: aluno_aj = st.selectbox("Aluno", nomes, key="aj_aluno")
        with col_aj2: tipo_aj = st.selectbox("Tipo", ["Bloqueio (ausência)","Trocar de SG","Trocar de local nessa semana"], key="aj_tipo")
        with col_aj3: sem_aj = st.number_input("Semana", min_value=1, max_value=int(num_semanas), key="aj_sem")
        if tipo_aj == "Trocar de SG":
            novo_sg_aj = st.number_input("Novo SG", min_value=1, max_value=sg_sugerido, key="aj_sg")
        if "ajustes" not in st.session_state: st.session_state.ajustes = []
        if st.button("Adicionar ajuste"):
            aj = {"aluno": aluno_aj, "tipo": tipo_aj, "semana": int(sem_aj)}
            if tipo_aj == "Trocar de SG": aj["novo_sg"] = int(novo_sg_aj)
            st.session_state.ajustes.append(aj)
            st.success("Ajuste adicionado!")
        for i, aj in enumerate(st.session_state.get("ajustes",[])):
            col_x, col_d = st.columns([1,10])
            with col_x:
                if st.button("❌", key=f"del_{i}"):
                    st.session_state.ajustes.pop(i); st.rerun()
            with col_d:
                st.write(f"**{aj['aluno']}** | {aj['tipo']} | Sem {aj['semana']}")

# ── GERAR ────────────────────────────────────────────────────────────────────
st.divider()
if st.button("🚀 Gerar Escala Completa", type="primary", use_container_width=True):
    if not especialidade: st.error("Preencha a especialidade!")
    elif not turma: st.error("Preencha a turma!")
    elif not alunos_finais: st.error("Adicione os alunos!")
    elif any(not l["nome"] for l in locais): st.error("Preencha todos os locais!")
    else:
        config = {
            "especialidade": especialidade, "ano_curso": ano_curso,
            "grupo": grupo_selecionado, "turma": turma,
            "data_inicio": str(data_inicio), "num_semanas": int(num_semanas),
            "num_sg": sg_sugerido, "locais": locais,
            "pares_bp": pares_bp, "rotacao_bp": rotacao_bp,
            "bloqueios": {"quinta_tarde": bloquear_quinta_tarde, "terca_parcial": bloquear_terca_parcial,
                          "feriados": [str(f) for f in feriados_finais]},
            "ajustes": st.session_state.get("ajustes", []),
            "alunos": alunos_finais,
        }
        with st.spinner("Gerando escala... ⏳"):
            try:
                resultado = gerar_escala(config)
                st.success("✅ Escala gerada!")
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    st.download_button("📥 Baixar Excel (.xlsx)", data=resultado["xlsx"],
                        file_name=f"Escala_{especialidade}_{grupo_selecionado}_{turma}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)
                with col_dl2:
                    st.download_button("📥 Baixar PowerPoint (.pptx)", data=resultado["pptx"],
                        file_name=f"Escala_{especialidade}_{grupo_selecionado}_{turma}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True)
            except Exception as e:
                st.error(f"Erro: {e}"); st.exception(e)
