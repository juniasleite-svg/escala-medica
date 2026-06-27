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

            st.subheader("📋 O que a IA entendeu:")
            st.info(dados.get("resumo", ""))

            col_r1, col_r2, col_r3 = st.columns(3)
            with col_r1:
                st.metric("Especialidade", dados.get("especialidade","?"))
                st.metric("Grupo / Turma", f"{dados.get('grupo','?')} / {dados.get('turma','?')}")
            with col_r2:
                st.metric("Nº de Semanas", dados.get("num_semanas","?"))
                st.metric("Nº de Subgrupos", dados.get("num_sg","?"))
            with col_r3:
                st.metric("Data de Início", dados.get("data_inicio","?"))
                locais_nomes = [l["nome"] for l in dados.get("locais",[])]
                st.metric("Locais", ", ".join(locais_nomes))

            if dados.get("subgrupos"):
                with st.expander("👥 Subgrupos identificados"):
                    for sg_info in dados["subgrupos"]:
                        st.write(f"**SG{sg_info['sg']}:** {', '.join(sg_info.get('alunos',[]))}")

            st.subheader("✏️ O que deseja modificar?")
            modificacoes = st.multiselect(
                "Selecione as modificações desejadas:",
                ["Trocar aluno de subgrupo", "Substituir aluno ausente", 
                 "Adicionar local", "Remover local", "Mudar turnos de um local",
                 "Alterar datas", "Mudar número de semanas", "Alterar bloqueios fixos"]
            )

            if "Trocar aluno de subgrupo" in modificacoes or "Substituir aluno ausente" in modificacoes:
                with st.expander("👤 Ajustes de alunos", expanded=True):
                    todos_alunos = []
                    for sg_info in dados.get("subgrupos",[]):
                        todos_alunos.extend(sg_info.get("alunos",[]))
                    
                    col_m1, col_m2, col_m3 = st.columns(3)
                    with col_m1:
                        aluno_sel = st.selectbox("Aluno", sorted(todos_alunos) if todos_alunos else ["(nenhum identificado)"])
                    with col_m2:
                        tipo_mod = st.selectbox("Tipo", ["Trocar de SG", "Marcar ausente na semana"])
                    with col_m3:
                        sem_mod = st.number_input("Semana", min_value=1, max_value=int(dados.get("num_semanas",8)), value=1)
                    
                    if tipo_mod == "Trocar de SG":
                        novo_sg_mod = st.number_input("Novo SG", min_value=1, max_value=int(dados.get("num_sg",6)), value=1)

                    if "ajustes_import" not in st.session_state:
                        st.session_state.ajustes_import = []
                    
                    if st.button("Adicionar ajuste"):
                        aj = {"aluno": aluno_sel, "tipo": tipo_mod, "semana": int(sem_mod)}
                        if tipo_mod == "Trocar de SG":
                            aj["novo_sg"] = int(novo_sg_mod)
                        st.session_state.ajustes_import.append(aj)
                        st.success("Ajuste adicionado!")

                    for aj in st.session_state.get("ajustes_import", []):
                        st.write(f"• **{aj['aluno']}** | {aj['tipo']} | Sem {aj['semana']}")

            if st.button("🚀 Gerar Escala Modificada", type="primary", use_container_width=True):
                st.info("🔄 Gerando escala com as modificações... Esta funcionalidade será ativada em breve!")

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
