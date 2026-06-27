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
        timeout=120
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
    # 3. Remove comentários // e tenta parsear
    linhas = [l for l in texto.split('\n') if not l.strip().startswith('//') and not l.strip().startswith('#')]
    txt = '\n'.join(linhas)
    inicio = txt.find('{')
    if inicio >= 0:
        try: return json.loads(txt[inicio:])
        except: pass
        depth = 0
        for i, ch in enumerate(txt[inicio:]):
            if ch == '{': depth += 1
            elif ch == '}':
                depth -= 1
                if depth == 0:
                    try: return json.loads(txt[inicio:inicio+i+1])
                    except: pass
    # 4. Entre primeiro { e último }
    try:
        s = texto.index('{'); e = texto.rindex('}') + 1
        return json.loads(texto[s:e])
    except: pass
    return None

# ── System prompt para geração ───────────────────────────────────────────────
SYSTEM_GERAR = """Você é um especialista em escalas de internato médico.
Gere a escala completa seguindo EXATAMENTE as regras do briefing.
Responda APENAS com JSON válido, sem texto antes ou depois, sem comentários //.

Estrutura obrigatória:
{
  "confirmacao": "resumo do que entendeu",
  "calendario_rodizio": [
    {"semana": 1, "periodo": "06/07-12/07", "alocacao": {"SG1": "Enf", "SG2": "PS"}}
  ],
  "escala_detalhada": [
    {"semana": 1, "data": "06/07", "dia": "Seg", "local": "Enf", "turno": "Manhã", "horario": "07-13h", "horas": 6, "sg": 1, "alunos": ["Nome1","Nome2"], "ra": ""}
  ],
  "resumo_horas": [
    {"sg": 1, "nome": "Nome", "ra": "", "total_horas": 40, "semanas": [40,40,40,40,40,40,40,40], "plantoes_enf_fds": 1, "plantoes_ps_fds": 4, "cinderelas": 0}
  ],
  "auditoria": {"ch_ok": true, "cobertura_ok": true, "erros": [], "avisos": [], "aprovado": true}
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
# MODO IMPORTAR
# ════════════════════════════════════════════════════════════════════════════
if modo == "📂 Importar escala existente e modificar":
    st.header("📂 Importar Escala Existente")
    st.caption("Faça upload de qualquer Excel de escala já feita — a IA lê e preenche todos os campos automaticamente.")

    col_up1, col_up2 = st.columns(2)
    with col_up1:
        arquivo_escala = st.file_uploader("Upload da escala existente (.xlsx)", type=["xlsx"], key="imp_escala")
    with col_up2:
        arquivo_alunos_imp = st.file_uploader("Base de alunos (opcional)", type=["xlsx"], key="imp_alunos_file")

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

                prompt = f"""Analise esta escala médica e extraia UM briefing completo em JSON.

CONTEÚDO DO EXCEL:
{conteudo}

Retorne APENAS JSON válido (sem markdown, sem comentários) com esta estrutura exata:
{{
  "especialidade": "",
  "grupo": "",
  "turma": "",
  "ano_curso": "4º Ano",
  "data_inicio": "YYYY-MM-DD",
  "num_semanas": 8,
  "num_sg": 6,
  "alunos_por_sg": {{"1": ["Nome1","Nome2"], "2": ["Nome3"]}},
  "locais": [
    {{
      "nome": "", "abrev": "", "obs": "",
      "manha": "07-13h", "min_manha": 0, "max_manha": 6,
      "tarde": "12-18h", "min_tarde": 3, "max_tarde": 4,
      "bloqueios_tarde": [{{"dia": "Qui", "tipo": "Sem tarde", "horario": ""}}, {{"dia": "Ter", "tipo": "Horário reduzido", "horario": "12-16h"}}],
      "cinderela": "", "min_cind": 0, "max_cind": 0, "dias_cind": [],
      "fds": false, "fds_quem": "", "fds_comp": "",
      "fds_manha": "", "fds_min_manha": 0, "fds_max_manha": 0,
      "fds_tarde": "", "fds_min_tarde": 0, "fds_max_tarde": 0,
      "servico2": "", "servico2_obs": ""
    }}
  ],
  "rodizio_desc": "descrição do rodízio",
  "regras_especiais": {{
    "quinta": "Sem tarde (ENAMED)",
    "terca": "Tarde encurtada 12-16h",
    "limite_ch": 40,
    "limite_abs": 43,
    "fds": "",
    "extras": ""
  }},
  "resumo": "descrição em português do que você entendeu"
}}"""

                resposta = chamar_claude([{"role": "user", "content": prompt}], max_tokens=6000)
                if resposta:
                    dados = extrair_json(resposta)
                    if dados:
                        st.session_state.escala_importada = dados
                        st.success("✅ Escala analisada! Formulário preenchido automaticamente.")
                    else:
                        st.error("Não consegui extrair o JSON. Tente novamente.")
                        st.text(resposta[:300])

    # Mostrar formulário pré-preenchido
    if "escala_importada" in st.session_state:
        dados = st.session_state.escala_importada
        st.info(f"💡 **Resumo:** {dados.get('resumo','')}")
        st.divider()

        # Identificação
        st.subheader("1️⃣ Identificação")
        col1, col2 = st.columns(2)
        with col1:
            esp_i = st.text_input("Especialidade", value=dados.get("especialidade",""), key="ii_esp")
            grupo_i = st.text_input("Grupo", value=dados.get("grupo",""), key="ii_grupo")
            ano_i = st.selectbox("Ano", ["3º Ano","4º Ano","5º Ano","6º Ano"],
                index=["3º Ano","4º Ano","5º Ano","6º Ano"].index(dados.get("ano_curso","4º Ano"))
                if dados.get("ano_curso") in ["3º Ano","4º Ano","5º Ano","6º Ano"] else 1, key="ii_ano")
        with col2:
            turma_i = st.text_input("Turma", value=dados.get("turma",""), key="ii_turma")
            try: d0 = datetime.date.fromisoformat(dados.get("data_inicio",""))
            except: d0 = datetime.date.today()
            data_i = st.date_input("Data início", value=d0, key="ii_data")
            nsem_i = st.number_input("Nº semanas", 1, 20, int(dados.get("num_semanas",8)), key="ii_nsem")

        # Alunos
        st.subheader("2️⃣ Alunos por Subgrupo")
        alunos_i = {}
        for sg, nomes in sorted(dados.get("alunos_por_sg",{}).items(), key=lambda x: int(x[0])):
            with st.expander(f"SG{sg} — {len(nomes)} alunos", expanded=False):
                txt = st.text_area(f"Alunos SG{sg}", value="\n".join(nomes), key=f"ii_sg_{sg}", height=100)
                alunos_i[sg] = [n.strip() for n in txt.strip().split("\n") if n.strip()]

        # Locais (resumido — editável)
        st.subheader("3️⃣ Locais")
        locais_i = dados.get("locais", [])
        locais_editados_i = []
        for idx, loc in enumerate(locais_i):
            with st.expander(f"📍 {loc.get('nome','Local')} — clique para editar", expanded=False):
                col_la, col_lb = st.columns(2)
                with col_la:
                    nome_e = st.text_input("Nome", value=loc.get("nome",""), key=f"ii_ln_{idx}")
                    manha_e = st.text_input("Manhã", value=loc.get("manha","07-13h"), key=f"ii_lm_{idx}")
                    tarde_e = st.text_input("Tarde normal", value=loc.get("tarde","12-18h"), key=f"ii_lt_{idx}")
                with col_lb:
                    cind_e = st.text_input("Cinderela (vazio=sem)", value=loc.get("cinderela",""), key=f"ii_lc_{idx}")
                    fds_e = st.checkbox("Tem FDS?", value=loc.get("fds",False), key=f"ii_lfds_{idx}")
                    obs_e = st.text_input("Obs", value=loc.get("obs",""), key=f"ii_lobs_{idx}")

                # Bloqueios
                bloqs = loc.get("bloqueios_tarde", [])
                bloqs_e = []
                st.caption("Bloqueios de tarde:")
                for bi, bloq in enumerate(bloqs):
                    bc1, bc2, bc3 = st.columns(3)
                    with bc1:
                        dias_op = ["Seg","Ter","Qua","Qui","Sex"]
                        idx_dia = dias_op.index(bloq.get("dia","Qui")) if bloq.get("dia") in dias_op else 3
                        d_e = st.selectbox("Dia", dias_op, index=idx_dia, key=f"ii_bd_{idx}_{bi}")
                    with bc2:
                        t_e = st.selectbox("Tipo", ["Sem tarde","Horário reduzido"],
                            index=0 if bloq.get("tipo","")=="Sem tarde" else 1, key=f"ii_bt_{idx}_{bi}")
                    with bc3:
                        h_e = st.text_input("Horário", value=bloq.get("horario","12-16h"), key=f"ii_bh_{idx}_{bi}")
                    bloqs_e.append({"dia": d_e, "tipo": t_e, "horario": h_e if t_e=="Horário reduzido" else ""})

                loc_e = dict(loc)
                loc_e.update({"nome": nome_e, "manha": manha_e, "tarde": tarde_e,
                               "cinderela": cind_e, "fds": fds_e, "obs": obs_e,
                               "bloqueios_tarde": bloqs_e})
                locais_editados_i.append(loc_e)

        # Rodízio
        st.subheader("4️⃣ Rodízio")
        rodizio_i = st.text_area("Tabela de rodízio",
            value=dados.get("rodizio_desc",""), height=100, key="ii_rod")

        # Regras
        st.subheader("5️⃣ Regras Especiais")
        reg = dados.get("regras_especiais", {})
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            quinta_i = st.text_input("Quinta-feira", value=reg.get("quinta","Sem tarde (ENAMED)"), key="ii_qui")
            terca_i = st.text_input("Terça-feira", value=reg.get("terca","Tarde 12-16h"), key="ii_ter")
        with col_r2:
            ch_i = st.number_input("Limite CH padrão", 20, 60, int(reg.get("limite_ch",40)), key="ii_ch")
            chabs_i = st.number_input("Limite CH absoluto", 20, 60, int(reg.get("limite_abs",43)), key="ii_chabs")
        fds_reg_i = st.text_input("Regras FDS", value=reg.get("fds",""), key="ii_fds")
        extras_i = st.text_area("Regras extras", value=reg.get("extras",""), height=60, key="ii_ext")

        # Modificações
        st.subheader("✏️ O que deseja modificar?")
        modif_i = st.text_area("Descreva livremente (deixe vazio para manter tudo igual)",
            height=80, key="ii_modif",
            placeholder="Ex: Trocar Ana do SG2 para SG4 nas semanas 5-8\nAdicionar feriado em 09/07\nPA deve ter FDS manhã e tarde")

        if st.button("🚀 Gerar Escala com IA", type="primary", use_container_width=True):
            briefing = f"""
ESPECIALIDADE: {esp_i} | GRUPO: {grupo_i} | TURMA: {turma_i} | ANO: {ano_i}
DATA INÍCIO: {data_i} | SEMANAS: {nsem_i}

ALUNOS:
{json.dumps(alunos_i, ensure_ascii=False, indent=2)}

LOCAIS:
{json.dumps(locais_editados_i, ensure_ascii=False, indent=2)}

RODÍZIO:
{rodizio_i}

REGRAS:
Quinta: {quinta_i}
Terça: {terca_i}
CH limite: {ch_i}h | Absoluto: {chabs_i}h
FDS: {fds_reg_i}
Extras: {extras_i}

MODIFICAÇÕES SOLICITADAS:
{modif_i if modif_i else "Nenhuma — manter a lógica original"}
"""
            st.session_state.briefing_atual = briefing
            st.session_state.config_atual = {
                "especialidade": esp_i, "grupo": grupo_i, "turma": turma_i,
                "ano_curso": ano_i, "data_inicio": str(data_i), "num_semanas": int(nsem_i),
                "locais": locais_editados_i, "alunos_por_sg": alunos_i,
                "rodizio_desc": rodizio_i,
                "regras_especiais": {"quinta": quinta_i, "terca": terca_i,
                    "limite_ch": int(ch_i), "limite_abs": int(chabs_i), "fds": fds_reg_i},
                "pares": [], "blocos": [],
            }
            with st.spinner("Gerando escala... ⏳"):
                resp = chamar_claude(
                    [{"role": "user", "content": f"Gere a escala:\n{briefing}"}],
                    system_prompt=SYSTEM_GERAR, max_tokens=8000
                )
                if resp:
                    st.session_state.escala_gerada = resp
                    st.session_state.esp_atual = esp_i
                    st.session_state.grupo_atual = grupo_i
                    st.session_state.turma_atual = turma_i
                    st.rerun()

    # Mostrar resultado se gerado
    if "escala_gerada" in st.session_state and "esp_atual" in st.session_state:
        st.divider()
        st.header("📊 Resultado")
        mostrar_resultado(
            st.session_state.escala_gerada,
            st.session_state.get("esp_atual",""),
            st.session_state.get("grupo_atual",""),
            st.session_state.get("turma_atual","")
        )

    st.stop()

# ════════════════════════════════════════════════════════════════════════════
# MODO CRIAR NOVA ESCALA
# ════════════════════════════════════════════════════════════════════════════
st.title("🏥 Gerador de Escalas Médicas")
st.header("📋 Briefing da Escala")
st.caption("Preencha com atenção — quanto mais detalhado, mais precisa a escala gerada.")

# BLOCO 1
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

# BLOCO 2
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
                    st.success(f"✅ {len(df_f)} alunos em {len(alunos_por_sg)} SGs")
        except Exception as e:
            st.error(f"Erro: {e}")
    if not alunos_por_sg:
        st.caption("Ou digite manualmente:")
        for sg in range(1, int(num_sg)+1):
            txt = st.text_area(f"SG{sg} (um nome por linha)", key=f"sg_{sg}", height=80)
            if txt.strip():
                alunos_por_sg[str(sg)] = [n.strip() for n in txt.strip().split("\n") if n.strip()]

# BLOCO 3
with st.expander("📍 Bloco 3 — Locais de Rodízio", expanded=True):
    num_locais = st.number_input("Número de locais", 2, 8, 3)
    locais = []
    for i in range(int(num_locais)):
        with st.container():
            st.markdown(f"### Local {i+1}")
            col_a, col_b = st.columns(2)
            with col_a:
                nome_l = st.text_input("Nome do local", key=f"ln_{i}", placeholder="ex: Enfermaria")
                abrev_l = st.text_input("Abreviação", key=f"la_{i}", placeholder="ex: Enf")
            with col_b:
                obs_l = st.text_input("Observações", key=f"lobs_{i}", placeholder="ex: FDS feito pelos alunos do Amb")
                unir = st.checkbox("🔗 Unir com outro serviço neste bloco", key=f"lunir_{i}")
                s2_nome = st.text_input("2º serviço (nome)", key=f"ln2_{i}") if unir else ""
                s2_obs = st.text_input("2º serviço (regras)", key=f"lobs2_{i}") if unir else ""

            # Duração por subgrupo
            st.markdown("**📅 Duração por Subgrupo neste local (semanas):**")
            st.caption("Defina quantas semanas cada SG fica neste local. Use 0 se o SG não passa por aqui.")
            n_sgs_atual = len(alunos_por_sg) if alunos_por_sg else int(num_sg)
            duracao_sgs = {}
            dur_cols = st.columns(min(n_sgs_atual, 6))
            for sg_idx in range(n_sgs_atual):
                sg_num = sg_idx + 1
                with dur_cols[sg_idx % 6]:
                    dur = st.number_input(f"SG{sg_num}", 0, int(num_semanas), 0,
                                          key=f"dur_{i}_{sg_num}",
                                          help=f"Semanas que SG{sg_num} fica em {nome_l or f'Local {i+1}'}")
                    if dur > 0:
                        duracao_sgs[str(sg_num)] = int(dur)
            if duracao_sgs:
                total_sem = sum(duracao_sgs.values())
                media = total_sem / len(duracao_sgs)
                st.caption(f"Total alocado: {total_sem} semanas × SG | Média: {media:.1f} sem/SG")


            st.markdown("**⏰ Turnos:**")
            tab_util, tab_fds = st.tabs(["📅 Dias Úteis", "🏖️ Final de Semana"])

            with tab_util:
                col_u1, col_u2, col_u3 = st.columns(3)
                with col_u1:
                    st.markdown("**🌅 Manhã**")
                    tem_m = st.checkbox("Tem manhã?", value=True, key=f"tm_{i}")
                    hor_m = st.text_input("Horário", value="07-13h", key=f"hm_{i}") if tem_m else ""
                    min_m = st.number_input("Mín/dia", 0, 20, 0, key=f"mnm_{i}") if tem_m else 0
                    max_m = st.number_input("Máx/dia", 0, 20, 6, key=f"mxm_{i}") if tem_m else 0

                with col_u2:
                    st.markdown("**🌇 Tarde**")
                    tem_t = st.checkbox("Tem tarde?", value=True, key=f"tt_{i}")
                    if tem_t:
                        hor_t = st.text_input("Horário normal", value="12-18h", key=f"ht_{i}")
                        min_t = st.number_input("Mín/dia", 0, 20, 3, key=f"mnt_{i}")
                        max_t = st.number_input("Máx/dia", 0, 20, 4, key=f"mxt_{i}")
                        st.markdown("**🚫 Bloqueios:**")
                        n_bloq = st.number_input("Quantos?", 0, 5, 1, key=f"nbloq_{i}")
                        bloqueios_t = []
                        for b in range(int(n_bloq)):
                            dias_op = ["Seg","Ter","Qua","Qui","Sex"]
                            bc1, bc2, bc3 = st.columns(3)
                            with bc1:
                                d_b = st.selectbox("Dia", dias_op, key=f"dbloq_{i}_{b}",
                                    index=3 if b==0 else (1 if b==1 else 0))
                            with bc2:
                                t_b = st.selectbox("Tipo", ["Sem tarde","Horário reduzido"], key=f"tbloq_{i}_{b}")
                            with bc3:
                                h_b = st.text_input("Horário", value="12-16h", key=f"hbloq_{i}_{b}") if t_b=="Horário reduzido" else ""
                            bloqueios_t.append({"dia": d_b, "tipo": t_b, "horario": h_b})
                    else:
                        hor_t, min_t, max_t, bloqueios_t = "", 0, 0, []

                with col_u3:
                    st.markdown("**🌙 Cinderela**")
                    tem_c = st.checkbox("Tem cinderela?", key=f"tc_{i}")
                    if tem_c:
                        hor_c = st.text_input("Horário", value="19-23h", key=f"hc_{i}")
                        min_c = st.number_input("Mín/dia", 0, 10, 0, key=f"mnc_{i}")
                        max_c = st.number_input("Máx/dia", 0, 10, 2, key=f"mxc_{i}")
                        dias_c = st.multiselect("Dias", ["Seg","Ter","Qua","Qui","Sex"], default=["Sex"], key=f"dc_{i}")
                    else:
                        hor_c, min_c, max_c, dias_c = "", 0, 0, []

            with tab_fds:
                tem_fds = st.checkbox("Tem plantão no FDS?", key=f"lfds_{i}")
                if tem_fds:
                    col_f1, col_f2, col_f3 = st.columns(3)
                    with col_f1:
                        st.markdown("**🌅 Manhã FDS**")
                        tem_fm = st.checkbox("Tem?", value=True, key=f"tfm_{i}")
                        hor_fm = st.text_input("Horário", value="07-12h", key=f"hfm_{i}") if tem_fm else ""
                        min_fm = st.number_input("Mín", 0, 10, 1, key=f"mnfm_{i}") if tem_fm else 0
                        max_fm = st.number_input("Máx", 0, 10, 2, key=f"mxfm_{i}") if tem_fm else 0
                    with col_f2:
                        st.markdown("**🌇 Tarde FDS**")
                        tem_ft = st.checkbox("Tem?", key=f"tft_{i}")
                        hor_ft = st.text_input("Horário", value="13-19h", key=f"hft_{i}") if tem_ft else ""
                        min_ft = st.number_input("Mín", 0, 10, 1, key=f"mnft_{i}") if tem_ft else 0
                        max_ft = st.number_input("Máx", 0, 10, 2, key=f"mxft_{i}") if tem_ft else 0
                    with col_f3:
                        st.markdown("**🌙 Cinderela FDS**")
                        tem_fc = st.checkbox("Tem?", key=f"tfc_{i}")
                        hor_fc = st.text_input("Horário", value="19-23h", key=f"hfc_{i}") if tem_fc else ""
                        min_fc = st.number_input("Mín", 0, 10, 0, key=f"mnfc_{i}") if tem_fc else 0
                        max_fc = st.number_input("Máx", 0, 10, 2, key=f"mxfc_{i}") if tem_fc else 0
                    quem_fds = st.text_input("Quem faz o FDS?", key=f"lfdsquem_{i}",
                        placeholder="ex: alunos do próprio local | alunos do Ambulatório")
                    comp_fds = st.text_input("Compensação?", key=f"lfdscomp_{i}",
                        placeholder="ex: perde 1 tarde | nenhuma")
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

# BLOCO 4 — Rodízio com sugestão automática
with st.expander("🔄 Bloco 4 — Tabela de Rodízio", expanded=True):

    def gerar_sugestoes(n_sg, n_locais, n_sem, nomes_loc):
        locs = nomes_loc if nomes_loc and all(nomes_loc) else [f"Local{j+1}" for j in range(n_locais)]
        sugestoes = []

        # Sugestão por pares
        if n_sg % n_locais == 0:
            tam = n_sg // n_locais
            sems_loc = n_sem // n_locais
            resto = n_sem % n_locais
            linhas = []
            for p in range(n_locais):
                sgs = "+".join([f"SG{s}" for s in range(p*tam+1, (p+1)*tam+1)])
                sem_at = 1
                rot = []
                for j in range(n_locais):
                    q = sems_loc + (1 if j < resto else 0)
                    rot.append(f"{locs[(p+j)%n_locais]} S{sem_at}-{sem_at+q-1}")
                    sem_at += q
                linhas.append(f"Par{p+1} = {sgs}: {' → '.join(rot)}")
            sugestoes.append({"nome": f"🔄 Por pares ({tam} SG/par, {sems_loc}{'–'+str(sems_loc+1) if resto else ''} sem/local)", "texto": "\n".join(linhas)})

        # Individual
        sems_loc2 = n_sem // n_locais
        resto2 = n_sem % n_locais
        linhas2 = []
        for sg in range(1, n_sg+1):
            sem_at = 1
            rot = []
            for j in range(n_locais):
                q = sems_loc2 + (1 if j < resto2 else 0)
                rot.append(f"{locs[(sg-1+j)%n_locais]} S{sem_at}-{sem_at+q-1}")
                sem_at += q
            linhas2.append(f"SG{sg}: {' → '.join(rot)}")
        sugestoes.append({"nome": "🔄 Individual (cada SG passa por todos os locais)", "texto": "\n".join(linhas2)})

        # Semanal estilo Cirurgia
        if n_sg >= n_locais:
            linhas3 = []
            for s in range(1, n_sem+1):
                alocacao = " | ".join([f"SG{((j+s-1)%n_sg)+1}→{locs[j%n_locais]}" for j in range(n_locais)])
                linhas3.append(f"Sem {s}: {alocacao}")
            sugestoes.append({"nome": "📅 Semanal (1 SG/local/semana, estilo Cirurgia)", "texto": "\n".join(linhas3)})

        return sugestoes

    nomes_loc_atual = [l.get("nome","") for l in locais]
    n_sg_atual = len(alunos_por_sg) if alunos_por_sg else int(num_sg)

    if st.button("💡 Sugerir rodízio automaticamente", use_container_width=True):
        st.session_state.sugestoes_rod = gerar_sugestoes(n_sg_atual, int(num_locais), int(num_semanas), nomes_loc_atual)

    if st.session_state.get("sugestoes_rod"):
        st.markdown("**Escolha uma opção:**")
        for idx, sug in enumerate(st.session_state.sugestoes_rod):
            with st.expander(sug["nome"], expanded=(idx==0)):
                st.code(sug["texto"])
                if st.button(f"✅ Usar esta", key=f"usar_{idx}"):
                    st.session_state.rodizio_escolhido = sug["texto"]
                    st.success("Aplicado!")

    rodizio_desc = st.text_area("Tabela de rodízio (edite à vontade)",
        value=st.session_state.get("rodizio_escolhido",""), height=150,
        placeholder="Ex:\nPar1 = SG1+SG2: Enf S1-3 → PS S4-5 → Amb S6-8\nPar2 = SG3+SG4: PS S1-3 → Amb S4-5 → Enf S6-8")

# BLOCO 5
with st.expander("⚙️ Bloco 5 — Regras Especiais", expanded=True):
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        regra_quinta = st.text_input("Quinta-feira", value="Sem tarde (ENAMED para todos)")
        regra_terca = st.text_input("Terça-feira", value="Tarde encurtada 12-16h (aula às 16h)")
        limite_ch = st.number_input("Limite CH padrão (h)", 20, 60, 40)
    with col_r2:
        limite_abs = st.number_input("Limite CH absoluto (h)", 20, 60, 43)
        regra_fds = st.text_area("Regras de plantão FDS", height=80,
            placeholder="ex: FDS PS = alunos do Amb, 3M+3T, comp: perde 1 tarde Amb")
        regras_extras = st.text_area("Outras regras", height=80,
            placeholder="ex: Ambulatório: manhã 08-12h")

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
        with st.spinner("IA gerando sua escala... ⏳ Pode levar até 2 minutos"):
            resposta = chamar_claude(
                [{"role": "user", "content": f"Gere a escala:\n{briefing}"}],
                system_prompt=SYSTEM_GERAR, max_tokens=8000
            )
            if resposta:
                st.session_state.escala_gerada = resposta
                st.session_state.esp_atual = especialidade
                st.session_state.grupo_atual = grupo
                st.session_state.turma_atual = turma
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
