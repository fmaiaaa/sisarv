# -*- coding: utf-8 -*-
"""
SisArv - Envio de inventário botânico via Streamlit
Login/senha configuráveis e upload de planilha (xlsx, csv, etc.).
Design de referência: Direcional (Simulador Imobiliário).
"""

import streamlit as st
import pandas as pd
import io
import sys
import os
import threading
import time

# Cores e estilo (referência Direcional)
COR_AZUL_ESC = "#002c5d"
COR_VERMELHO = "#e30613"
COR_FUNDO = "#fcfdfe"
COR_BORDA = "#eef2f6"
COR_TEXTO_MUTED = "#64748b"
COR_INPUT_BG = "#f0f2f6"

# Importa o módulo ws (mesmo diretório)
try:
    import ws
    from ws import run_sisarv, preprocessar_df
except ImportError:
    ws = None
    run_sisarv = None
    preprocessar_df = None


def aplicar_estilo():
    st.markdown(f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800;900&family=Inter:wght@300;400;500;600;700&display=swap');

        html, body, [data-testid="stAppViewContainer"] {{
            font-family: 'Inter', sans-serif;
            color: {COR_AZUL_ESC};
            background-color: {COR_FUNDO};
        }}

        h1, h2, h3, h4 {{
            font-family: 'Montserrat', sans-serif !important;
            color: {COR_AZUL_ESC} !important;
            font-weight: 800;
            text-align: center;
        }}

        .block-container {{ max-width: 900px !important; padding: 2rem !important; }}

        div[data-baseweb="input"] {{
            border-radius: 8px !important;
            border: 1px solid #e2e8f0 !important;
            background-color: {COR_INPUT_BG} !important;
        }}

        .stButton {{
            display: flex !important;
            justify-content: center !important;
            width: 100% !important;
        }}

        .stButton button {{
            font-family: 'Inter', sans-serif;
            border-radius: 8px !important;
            padding: 0 20px !important;
            width: 100% !important;
            height: 38px !important;
            min-height: 38px !important;
            font-weight: 700 !important;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }}

        .stButton button[kind="primary"] {{
            background: {COR_VERMELHO} !important;
            color: #ffffff !important;
            border: none !important;
        }}

        .stButton button[kind="primary"]:hover {{
            background: #c40510 !important;
        }}

        .header-container {{
            text-align: center;
            padding: 40px 0;
            background: #ffffff;
            margin-bottom: 40px;
            border-radius: 0 0 24px 24px;
            border-bottom: 1px solid {COR_BORDA};
            box-shadow: 0 10px 25px -15px rgba(0,44,93,0.15);
        }}

        .header-title {{
            font-family: 'Montserrat', sans-serif;
            color: {COR_AZUL_ESC};
            font-size: 2rem;
            font-weight: 900;
            margin: 0;
            text-transform: uppercase;
            letter-spacing: 0.15em;
        }}

        .header-subtitle {{
            color: {COR_AZUL_ESC};
            font-size: 0.95rem;
            font-weight: 600;
            margin-top: 10px;
            opacity: 0.85;
        }}

        .card {{
            background: #ffffff;
            padding: 24px;
            border-radius: 16px;
            border: 1px solid {COR_BORDA};
            margin-bottom: 24px;
        }}

        .footer {{ text-align: center; padding: 40px 0; color: {COR_AZUL_ESC}; font-size: 0.8rem; opacity: 0.7; }}
        </style>
    """, unsafe_allow_html=True)


def carregar_planilha(uploaded_file):
    """Lê xlsx, csv ou similar e retorna DataFrame."""
    nome = (uploaded_file.name or "").lower()
    raw = uploaded_file.read()
    if nome.endswith(".xlsx") or nome.endswith(".xls"):
        return pd.read_excel(io.BytesIO(raw))
    if nome.endswith(".csv"):
        try:
            return pd.read_csv(io.BytesIO(raw), encoding="utf-8", sep=";")
        except Exception:
            return pd.read_csv(io.BytesIO(raw), encoding="utf-8", sep=",")
    if nome.endswith(".ods"):
        try:
            return pd.read_excel(io.BytesIO(raw), engine="odf")
        except Exception:
            st.error("Para arquivos .ods instale: pip install odfpy")
            return None
    st.warning("Formato não suportado. Use .xlsx, .csv ou .ods.")
    return None


def main():
    st.set_page_config(page_title="SisArv - Inventário Botânico", page_icon="🌳", layout="centered")
    aplicar_estilo()

    st.markdown(
        '<div class="header-container">'
        '<div class="header-title">SisArv</div>'
        '<div class="header-subtitle">Envio de inventário botânico ao sistema</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    if run_sisarv is None or preprocessar_df is None:
        st.error("Módulo **ws.py** não encontrado. Coloque **sisarv_streamlit.py** na mesma pasta que **ws.py** e execute: `streamlit run sisarv_streamlit.py`")
        st.markdown('<div class="footer">Direcional Engenharia</div>', unsafe_allow_html=True)
        return

    with st.form("form_sisarv"):
        st.markdown("#### Credenciais do SisArv")
        login = st.text_input("E-mail (login)", placeholder="seu@email.com", key="login")
        senha = st.text_input("Senha", type="password", placeholder="••••••••", key="senha")
        st.markdown("---")
        st.markdown("#### Planilha de dados")
        uploaded = st.file_uploader(
            "Envie a planilha (XLSX, CSV ou ODS) com as colunas do inventário (Nº, Nome Vulgar, Nome Científico, etc.)",
            type=["xlsx", "xls", "csv", "ods"],
            key="upload",
        )
        enviar = st.form_submit_button("ENVIAR DADOS AO SISARV", type="primary")

    # Estado da execução em background
    if "sisarv_running" not in st.session_state:
        st.session_state.sisarv_running = False
    if "sisarv_stop_requested" not in st.session_state:
        st.session_state.sisarv_stop_requested = False
    if "sisarv_logs" not in st.session_state:
        st.session_state.sisarv_logs = []
    if "sisarv_result" not in st.session_state:
        st.session_state.sisarv_result = None
    if "sisarv_progress_current" not in st.session_state:
        st.session_state.sisarv_progress_current = 0
    if "sisarv_progress_total" not in st.session_state:
        st.session_state.sisarv_progress_total = 0

    # Se está rodando, mostrar progresso (tqdm-like) + log + botão Stop e atualizar a página periodicamente
    if st.session_state.sisarv_running:
        total = st.session_state.get("sisarv_progress_total", 0)
        current = st.session_state.get("sisarv_progress_current", 0)
        if total > 0:
            st.progress(current / total, text=f"Árvore **{current}** de **{total}**")
        else:
            st.caption("Aguardando início do preenchimento...")
        st.markdown("#### Log de execução")
        log_text = "\n".join(st.session_state.sisarv_logs[-50:]) if st.session_state.sisarv_logs else "(aguardando...)"
        st.code(log_text, language=None)
        stop_clicked = st.button("⏹ PARAR", type="secondary")
        if stop_clicked:
            st.session_state.sisarv_stop_requested = True
            st.rerun()
        time.sleep(1)
        st.rerun()

    # Se terminou (resultado disponível), mostrar e limpar
    if st.session_state.sisarv_result is not None:
        sucesso, arvores_nao_encontradas, erro = st.session_state.sisarv_result
        st.session_state.sisarv_result = None
        st.session_state.sisarv_running = False
        st.session_state.sisarv_stop_requested = False
        if erro:
            st.error(f"**Erro:** {erro}")
        elif sucesso:
            st.success("Processamento concluído.")
            if arvores_nao_encontradas:
                st.markdown("#### Árvores não encontradas nos selects")
                for n, vulg, cien in arvores_nao_encontradas:
                    st.caption(f"Nº {n}: {vulg!r} / {cien!r}")
                st.info(f"Total: **{len(arvores_nao_encontradas)}** árvore(s) não encontrada(s).")
        else:
            st.warning("Processamento finalizado com avisos. Veja o log acima.")
        st.markdown('<div class="footer">Direcional Engenharia | SisArv Inventário Botânico</div>', unsafe_allow_html=True)
        return

    if not enviar:
        st.markdown('<div class="footer">Informe login, senha e envie a planilha para continuar.</div>', unsafe_allow_html=True)
        return

    if not login or not senha:
        st.error("Preencha **e-mail** e **senha** do SisArv.")
        return

    if uploaded is None:
        st.error("Envie um arquivo (XLSX, CSV ou ODS).")
        return

    df_raw = carregar_planilha(uploaded)
    if df_raw is None or df_raw.empty:
        st.error("Não foi possível ler a planilha ou ela está vazia.")
        return

    df = preprocessar_df(df_raw)
    if df.empty:
        st.warning("Após o pré-processamento a planilha ficou vazia.")
        return

    st.success(f"Planilha carregada: **{len(df)}** linha(s).")
    with st.expander("Visualizar primeiras linhas"):
        st.dataframe(df.head(20), use_container_width=True, hide_index=True)

    def progress_callback(msg):
        st.session_state.sisarv_logs.append(msg)

    def progress_range_callback(atual, total):
        st.session_state.sisarv_progress_current = atual
        st.session_state.sisarv_progress_total = total

    def run_in_thread():
        try:
            result = run_sisarv(
                login.strip(),
                senha.strip(),
                df,
                progress_callback=progress_callback,
                should_stop=lambda: st.session_state.get("sisarv_stop_requested", False),
                progress_range_callback=progress_range_callback,
            )
            st.session_state.sisarv_result = result
        except Exception as e:
            st.session_state.sisarv_result = (False, [], str(e))
        finally:
            st.session_state.sisarv_running = False

    st.session_state.sisarv_logs = []
    st.session_state.sisarv_stop_requested = False
    st.session_state.sisarv_running = True
    st.session_state.sisarv_progress_current = 0
    st.session_state.sisarv_progress_total = len(df)
    thread = threading.Thread(target=run_in_thread)
    thread.start()

    # Redesenha a página em 1s para entrar no bloco "sisarv_running" (log + botão PARAR)
    st.markdown("#### Log de execução")
    st.code("(iniciando...)", language=None)
    time.sleep(1)
    st.rerun()


if __name__ == "__main__":
    main()
