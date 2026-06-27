import streamlit as st
import pandas as pd
from gerador_escala import gerar_escala
import io

st.set_page_config(page_title="Gerador de Escalas Médicas", page_icon="🏥", layout="wide")

st.title("🏥 Gerador de Escalas Médicas")

# ── PASSO 0: Upload do Excel de alunos ──────────────────────────────────────
st.header("Passo 0 — Base de Alunos")
arquivo_alunos = st.file_uploader(
    "Faça upload do Excel com grupos e subgrupos (TODOS_GRUPOS_SUBGRUPOS...xlsx)",
    type=["xlsx"],
    help="O arquivo deve ter abas por grupo (GRUPO A, GRUPO B...) com colunas: Nome Completo, RA, Grupo, OPÇÃO, Sub Grupo"
)

df_alunos = None
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
    ano_curso = st.selectbox("Ano do curso *", ["3º Ano", "4º Ano", "5º Ano", "6º Ano"])
with col2:
    grupo_opcoes = grupos_disponiveis if grupos_disponiveis else ["Grupo A", "Grupo B", "Grupo C", "Grupo D", "Grupo E", "Grupo F"]
    grupo_selecionado = st.selectbox("Grupo *", grupo_opcoes)
    turma = st.text_input("Turma *", placeholder="ex: T6")

col3, col4 = st.columns(2)
with col3:
    data_inicio = st.date_input("Data de início (segunda-feira) *")
with col4:
    num_semanas = st.number_input("Número de semanas *", min_value=1, max_value=20, value=8)

# ── PASSO 2: Locais de Rodízio ───────────────────────────────────────────────
st.header("Passo 2 — Locais de Rodízio")

num_locais = st.number_input("Quantos locais de estágio?", min_value=2, max_value=6, value=4)

# Lógica automática de SG
def calcular_sg_ideal(n_locais):
    # Regra: múltiplo de n_locais que dê grupos de 3-6 alunos
    opcoes_sg = {2: [4, 6, 8], 3: [6, 3], 4: [4, 8], 5: [5], 6: [6]}
    return opcoes_sg.get(n_locais, [4])[0]

sg_sugerido = calcular_sg_ideal(num_locais)
st.info(f"💡 Com {num_locais} locais, o sistema sugere usar **{sg_sugerido} subgrupos** para rotação equilibrada.")

locais = []
st.subheader("Configure cada local:")
for i in range(int(num_locais)):
    with st.expander(f"Local {i+1}", expanded=(i==0)):
        col_a, col_b = st.columns(2)
        with col_a:
            nome_local = st.text_input(f"Nome do local {i+1}", key=f"local_nome_{i}", placeholder="ex: UPA Mandic")
            tem_fds = st.checkbox("Funciona no fim de semana?", key=f"fds_{i}")
        with col_b:
            turno_m = st.checkbox("Turno Manhã (7-13h)", value=True, key=f"m_{i}")
            turno_t = st.checkbox("Turno Tarde (13-19h)", value=True, key=f"t_{i}")
            turno_c = st.checkbox("Turno Cinderela (18-23h)", key=f"c_{i}")
        
        col_c, col_d, col_e = st.columns(3)
        with col_c:
            vagas_m = st.number_input("Vagas manhã", min_value=0, max_value=20, value=6, key=f"vm_{i}") if turno_m else 0
        with col_d:
            vagas_t = st.number_input("Vagas tarde", min_value=0, max_value=20, value=4, key=f"vt_{i}") if turno_t else 0
        with col_e:
            vagas_c = st.number_input("Vagas cinderela", min_value=0, max_value=10, value=2, key=f"vc_{i}") if turno_c else 0

        locais.append({
            "nome": nome_local,
            "fds": tem_fds,
            "turno_m": turno_m, "vagas_m": vagas_m,
            "turno_t": turno_t, "vagas_t": vagas_t,
            "turno_c": turno_c, "vagas_c": vagas_c,
        })

# ── PASSO 3: Bloqueios ───────────────────────────────────────────────────────
st.header("Passo 3 — Bloqueios Fixos Semanais")

col_b1, col_b2 = st.columns(2)
with col_b1:
    bloquear_quinta_tarde = st.checkbox("Quinta-feira tarde bloqueada (ex: ENAMED 13-18h)", value=True)
with col_b2:
    bloquear_terca_parcial = st.checkbox("Terça-feira: campo só até 16h (aula 16-18h)", value=True)

bloqueios_extras = st.text_area(
    "Outros bloqueios (opcional)",
    placeholder="ex: Segunda 8-10h = Reunião clínica\nSexta 17h = Encerramento"
)

# ── PASSO 4: Alunos ──────────────────────────────────────────────────────────
st.header("Passo 4 — Alunos")

alunos_finais = []

if arquivo_alunos and grupo_selecionado:
    try:
        df_grupo = pd.read_excel(arquivo_alunos, sheet_name=grupo_selecionado)
        
        # Filtrar pela opção de SG correta
        opcao_col = "OPÇÃO"
        opcoes_disponiveis = df_grupo[opcao_col].unique() if opcao_col in df_grupo.columns else []
        
        opcao_escolhida = None
        for op in opcoes_disponiveis:
            if f"{sg_sugerido} SG" in str(op):
                opcao_escolhida = op
                break
        
        if opcao_escolhida is None and len(opcoes_disponiveis) > 0:
            opcao_escolhida = opcoes_disponiveis[0]

        if opcao_escolhida:
            df_filtrado = df_grupo[df_grupo[opcao_col] == opcao_escolhida].copy()
            st.success(f"✅ **{len(df_filtrado)} alunos** carregados do {grupo_selecionado} com {opcao_escolhida}")
            
            # Mostrar tabela
            df_display = df_filtrado[["Nome Completo", "RA", "Sub Grupo"]].copy()
            df_display["Sub Grupo"] = df_display["Sub Grupo"].astype(int)
            df_display = df_display.sort_values("Sub Grupo")
            st.dataframe(df_display, use_container_width=True, hide_index=True)
            
            alunos_finais = df_filtrado[["Nome Completo", "RA", "Sub Grupo"]].to_dict("records")
    except Exception as e:
        st.error(f"Erro ao carregar alunos: {e}")
else:
    st.info("👆 Faça o upload do Excel no Passo 0 para carregar os alunos automaticamente.")
    st.subheader("Ou adicione alunos manualmente:")
    
    num_alunos = st.number_input("Número de alunos", min_value=1, max_value=50, value=22)
    for i in range(int(num_alunos)):
        col_n, col_r, col_s = st.columns([3,2,1])
        with col_n:
            nome = st.text_input(f"Nome {i+1}", key=f"nome_{i}")
        with col_r:
            ra = st.text_input(f"RA {i+1}", key=f"ra_{i}")
        with col_s:
            sg = st.number_input(f"SG {i+1}", min_value=1, max_value=sg_sugerido, value=1, key=f"sg_{i}")
        if nome:
            alunos_finais.append({"Nome Completo": nome, "RA": ra, "Sub Grupo": sg})

# ── GERAR ────────────────────────────────────────────────────────────────────
st.divider()
if st.button("🚀 Gerar Escala Completa", type="primary", use_container_width=True):
    if not especialidade:
        st.error("Preencha a especialidade!")
    elif not turma:
        st.error("Preencha a turma!")
    elif not alunos_finais:
        st.error("Adicione os alunos!")
    elif any(not l["nome"] for l in locais):
        st.error("Preencha o nome de todos os locais!")
    else:
        config = {
            "especialidade": especialidade,
            "ano_curso": ano_curso,
            "grupo": grupo_selecionado,
            "turma": turma,
            "data_inicio": str(data_inicio),
            "num_semanas": int(num_semanas),
            "num_sg": sg_sugerido,
            "locais": locais,
            "bloqueios": {
                "quinta_tarde": bloquear_quinta_tarde,
                "terca_parcial": bloquear_terca_parcial,
                "extras": bloqueios_extras
            },
            "alunos": alunos_finais
        }

        with st.spinner("Gerando escala... ⏳"):
            try:
                resultado = gerar_escala(config)
                st.success("✅ Escala gerada com sucesso!")

                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    st.download_button(
                        "📥 Baixar Excel (.xlsx)",
                        data=resultado["xlsx"],
                        file_name=f"Escala_{especialidade}_{grupo_selecionado}_{turma}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                with col_dl2:
                    st.download_button(
                        "📥 Baixar PowerPoint (.pptx)",
                        data=resultado["pptx"],
                        file_name=f"Escala_{especialidade}_{grupo_selecionado}_{turma}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )
            except Exception as e:
                st.error(f"Erro ao gerar escala: {e}")
                st.exception(e)
