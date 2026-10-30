import streamlit as st
import subprocess
import sys
import os
import shutil
import tempfile
import json
import html
import datetime

import drive_utils

# =====================================================================
# CONFIGURAÇÃO DA PÁGINA (CENTRAL BX DE AUTOMAÇÕES)
# =====================================================================
st.set_page_config(
    page_title="Central BX de Automações | Bx finance",
    page_icon="💠",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =====================================================================
# ESTILO CUSTOMIZADO (CSS) - FIEL AO LAYOUT DA IMAGEM DE REFERÊNCIA
# =====================================================================
CUSTOM_CSS = '''
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700;800;900&family=JetBrains+Mono:wght@400;600;700&display=swap');

/* Fundo Geral Escuro com Glow Roxo e Ondas Vetoriais no Rodapé */
.stApp {
    background-color: #07060e;
    background-image:
        radial-gradient(circle at 28% 22%, rgba(72, 56, 145, 0.28) 0%, rgba(15, 12, 30, 0.0) 52%),
        radial-gradient(circle at 75% 65%, rgba(62, 48, 128, 0.18) 0%, rgba(7, 6, 14, 0.0) 55%),
        linear-gradient(180deg, #0b0918 0%, #07060e 100%);
    color: #e2e8f0;
    font-family: 'Inter', sans-serif;
}

/* Ondas sutis decorativas no fundo inferior (idêntico ao mockup) */
.bx-bg-waves {
    position: fixed;
    bottom: 0;
    left: 0;
    width: 100%;
    height: 180px;
    pointer-events: none;
    z-index: 0;
    opacity: 0.45;
}

/* Ocultar elementos padrão do Streamlit */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}

/* Container Principal */
.main .block-container {
    padding-top: 2.8rem;
    padding-bottom: 3rem;
    max-width: 1280px;
    position: relative;
    z-index: 1;
}

/* =====================================================================
   HEADER: ESQUERDA "CENTRAL BX DE AUTOMAÇÕES" | DIREITA "Bx finance"
   ===================================================================== */
.bx-header-container {
    padding: 0.5rem 0 1.8rem 0;
    margin-bottom: 1.8rem;
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
}
.bx-title-wrapper {
    display: flex;
    flex-direction: column;
}
.bx-central-title {
    font-size: 3.1rem;
    font-weight: 900;
    line-height: 1.05;
    letter-spacing: -1px;
    color: #ffffff;
    font-family: 'Inter', sans-serif;
    text-transform: uppercase;
}
.bx-central-title .bx-accent {
    color: #7D75BE;
}
.bx-subtitle-row {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-top: 14px;
}
.bx-subtitle-dash {
    width: 22px;
    height: 2px;
    background-color: #7D75BE;
    display: inline-block;
}
.bx-subtitle-text {
    color: #a0aec0;
    font-size: 0.82rem;
    font-weight: 600;
    letter-spacing: 2.5px;
    text-transform: uppercase;
    font-family: 'JetBrains Mono', monospace;
}

/* Logo BX Finance no canto superior direito */
.bx-right-logo-wrapper {
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    justify-content: flex-start;
    text-align: right;
}
.bx-logo-text {
    font-size: 3.6rem;
    font-weight: 800;
    line-height: 0.9;
    letter-spacing: -3px;
    font-family: 'Inter', sans-serif;
}
.bx-logo-text .b { color: #ffffff; }
.bx-logo-text .x { color: #7D75BE; }
.bx-logo-sub {
    font-size: 1.15rem;
    font-weight: 400;
    color: #cbd5e0;
    letter-spacing: 1px;
    margin-top: 4px;
}
.bx-build-tag {
    color: #718096;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.72rem;
    letter-spacing: 1.5px;
    margin-top: 6px;
}

/* Título "SELECIONE O PROJETO" */
.bx-select-title {
    color: #ffffff;
    font-family: 'Inter', sans-serif;
    font-weight: 800;
    font-size: 1.15rem;
    text-transform: uppercase;
    letter-spacing: 3px;
    margin-bottom: 14px;
}

/* =====================================================================
   BOTÕES PILL DE SELEÇÃO DE PROJETO (SECONDARY = INATIVO, PRIMARY = ATIVO)
   ===================================================================== */
div[data-testid="stButton"] > button[kind="primary"] {
    background: linear-gradient(135deg, #7468C5 0%, #5B50A0 100%) !important;
    color: #ffffff !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 800 !important;
    font-size: 0.9rem !important;
    border: 1px solid rgba(168, 158, 240, 0.45) !important;
    border-radius: 999px !important;
    padding: 1.25rem 1.8rem !important;
    text-transform: uppercase !important;
    letter-spacing: 1.8px !important;
    box-shadow: 0 0 28px rgba(116, 104, 197, 0.55) !important;
    transition: all 0.3s ease !important;
}
div[data-testid="stButton"] > button[kind="primary"]:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 0 36px rgba(125, 117, 190, 0.75) !important;
}

div[data-testid="stButton"] > button[kind="secondary"] {
    background: rgba(18, 16, 32, 0.75) !important;
    color: #e2e8f0 !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 700 !important;
    font-size: 0.9rem !important;
    border: 1px solid rgba(125, 117, 190, 0.32) !important;
    border-radius: 999px !important;
    padding: 1.25rem 1.8rem !important;
    text-transform: uppercase !important;
    letter-spacing: 1.8px !important;
    box-shadow: none !important;
    transition: all 0.3s ease !important;
}
div[data-testid="stButton"] > button[kind="secondary"]:hover {
    border-color: #7D75BE !important;
    background: rgba(40, 35, 72, 0.55) !important;
    color: #ffffff !important;
    transform: translateY(-2px) !important;
}

/* Inputs de Diretório Premium */
.stTextInput > div > div > input {
    background-color: rgba(15, 14, 26, 0.85) !important;
    color: #ffffff !important;
    border: 1px solid rgba(125, 117, 190, 0.25) !important;
    border-radius: 12px !important;
    padding: 0.95rem 1.1rem !important;
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 0.92rem !important;
    transition: all 0.3s ease;
    box-shadow: 0 4px 12px rgba(0,0,0,0.35);
}
.stTextInput > div > div > input:focus {
    border-color: #7D75BE !important;
    box-shadow: 0 0 0 1px #7D75BE, 0 0 18px rgba(125, 117, 190, 0.25) !important;
}
.stTextInput label {
    color: #a0aec0 !important;
    font-weight: 600 !important;
    font-size: 0.82rem !important;
    text-transform: uppercase;
    letter-spacing: 1.2px;
    margin-bottom: 6px;
}

/* Divisor sutil entre os módulos */
.bx-module-divider {
    border-top: 1px solid rgba(125, 117, 190, 0.16);
    margin: 1.6rem 0 1.4rem 0;
}

/* =====================================================================
   CARDS DE STATUS À DIREITA (IDÊNTICOS À IMAGEM DE REFERÊNCIA)
   ===================================================================== */
.dash-card {
    background: rgba(20, 18, 38, 0.62);
    border: 1px solid rgba(125, 117, 190, 0.22);
    border-radius: 18px;
    padding: 26px 20px;
    text-align: center;
    position: relative;
    overflow: hidden;
    backdrop-filter: blur(8px);
    box-shadow: 0 8px 28px rgba(0, 0, 0, 0.45);
}
.dash-value {
    font-size: 2.3rem;
    font-weight: 800;
    color: #ffffff;
    line-height: 1.15;
    font-family: 'JetBrains Mono', monospace;
}
.dash-label {
    font-size: 0.72rem;
    color: #a0aec0;
    text-transform: uppercase;
    letter-spacing: 2px;
    font-weight: 600;
    margin-top: 8px;
}

/* Console Estilo Terminal Hacker/Finance */
.bx-terminal-wrapper {
    background-color: #0a0a0f;
    border: 1px solid rgba(125, 117, 190, 0.25);
    border-radius: 14px;
    box-shadow: inset 0 2px 20px rgba(0,0,0,0.5), 0 10px 30px rgba(0,0,0,0.8);
    overflow: hidden;
    margin-top: 1rem;
}
.bx-terminal-header {
    background-color: #12111d;
    padding: 12px 20px;
    border-bottom: 1px solid rgba(125, 117, 190, 0.18);
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.terminal-tab {
    color: #718096;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.8rem;
    display: flex;
    align-items: center;
    gap: 8px;
}
.terminal-dot {
    width: 12px; height: 12px; border-radius: 50%; display: inline-block;
}
.dot-red { background-color: #ff5f56; }
.dot-yellow { background-color: #ffbd2e; }
.dot-green { background-color: #27c93f; }

.live-badge {
    display: flex;
    align-items: center;
    color: #7D75BE;
    font-size: 0.75rem;
    font-weight: 800;
    font-family: 'Inter', sans-serif;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    background: rgba(98, 89, 164, 0.12);
    padding: 4px 12px;
    border-radius: 20px;
    border: 1px solid rgba(98, 89, 164, 0.25);
}
.pulse {
    width: 8px;
    height: 8px;
    background-color: #6259A4;
    border-radius: 50%;
    margin-right: 8px;
    box-shadow: 0 0 0 0 rgba(98, 89, 164, 1);
    animation: pulse-animation 2s infinite;
}
@keyframes pulse-animation {
    0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(98, 89, 164, 0.7); }
    70% { transform: scale(1); box-shadow: 0 0 0 8px rgba(98, 89, 164, 0); }
    100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(98, 89, 164, 0); }
}
.bx-terminal-body {
    padding: 20px;
    height: 480px;
    overflow-y: auto;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.85rem;
    line-height: 1.7;
    color: #a0aec0;
    display: flex;
    flex-direction: column;
}
.log-line { margin: 0; padding: 3px 0; border-bottom: 1px dashed rgba(255,255,255,0.03); }
.log-highlight { color: #7D75BE; font-weight: 700; }
.log-success { color: #38A169; font-weight: 700; text-shadow: 0 0 8px rgba(56,161,105,0.3); }
.log-error { color: #E53E3E; font-weight: 700; text-shadow: 0 0 8px rgba(229,62,62,0.3); }
.log-info { color: #63B3ED; }
.log-system { color: #CBD5E0; font-weight: 600; }
.log-time { color: #4A5568; margin-right: 10px; font-size: 0.75rem;}

/* Scrollbar Hacker */
::-webkit-scrollbar { width: 8px; height: 8px; }
::-webkit-scrollbar-track { background: #0a0a0c; }
::-webkit-scrollbar-thumb { background: #2d3748; border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: #6259A4; }
</style>

<!-- Ondas decorativas no rodapé idênticas ao mockup -->
<svg class="bx-bg-waves" viewBox="0 0 1440 200" fill="none" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="none">
    <path d="M0 110 C 320 40, 580 180, 920 110 C 1180 55, 1340 130, 1440 110" stroke="rgba(125, 117, 190, 0.22)" stroke-width="1.2" fill="none"/>
    <path d="M0 145 C 360 75, 620 205, 960 135 C 1200 85, 1360 155, 1440 140" stroke="rgba(125, 117, 190, 0.14)" stroke-width="1.2" fill="none"/>
</svg>
'''
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# =====================================================================
# DIRETÓRIOS DOS ROBÔS (ORGANIZADOS POR SUBPASTAS EM robos/)
# =====================================================================
CAMINHO_ARQUIVO_ULTIMO_USO = "ultimo_caminho.json"
PASTA_BASE_ROBOS = os.path.join(os.path.dirname(__file__), "robos")

# Scripts Grupo Cativa (robos/CATIVA/)
PASTA_ROBOS_CATIVA = os.path.join(PASTA_BASE_ROBOS, "CATIVA")
SCRIPT_CATIVA_PDF = os.path.join(PASTA_ROBOS_CATIVA, "pdf_pt1.py")
SCRIPT_CATIVA_SICOOB = os.path.join(PASTA_ROBOS_CATIVA, "sicoob.py")
SCRIPT_CATIVA_OUTROS = os.path.join(PASTA_ROBOS_CATIVA, "outros.py")

# Scripts Revpack (robos/REVPACK/)
PASTA_ROBOS_REVPACK = os.path.join(PASTA_BASE_ROBOS, "REVPACK")
SCRIPT_REVPACK_CB = os.path.join(PASTA_ROBOS_REVPACK, "REV - C.B. - AUT.py")
SCRIPT_REVPACK_CS = os.path.join(PASTA_ROBOS_REVPACK, "REVPACK - C.S..py")

# Caminhos Pré-Preenchidos Padrão
CAMINHOS_PADRAO = {
    "cativa_sicoob": r"2026\10 - Outubro\28.10.2026\SICOOB",
    "revpack_cb": r"EXTRATOS BANCÁRIOS\2026\10. Outubro\28",
    "revpack_cs": r"POSIÇÕES FUNDOS\2026\10 - OUT\Fundos 29.10\Trabalhados",
}

# =====================================================================
# FUNÇÕES DE APOIO E GERENCIAMENTO DE CAMINHOS
# =====================================================================
def carregar_caminhos_salvos():
    caminhos = dict(CAMINHOS_PADRAO)
    if os.path.exists(CAMINHO_ARQUIVO_ULTIMO_USO):
        try:
            with open(CAMINHO_ARQUIVO_ULTIMO_USO, "r", encoding="utf-8") as f:
                dados = json.load(f)
                if "caminho" in dados and "cativa_sicoob" not in dados:
                    caminhos["cativa_sicoob"] = dados["caminho"]
                for k in CAMINHOS_PADRAO:
                    if dados.get(k):
                        caminhos[k] = dados[k]
        except Exception:
            pass
    return caminhos


def salvar_caminho_especifico(chave, valor):
    caminhos = carregar_caminhos_salvos()
    caminhos[chave] = valor
    try:
        with open(CAMINHO_ARQUIVO_ULTIMO_USO, "w", encoding="utf-8") as f:
            json.dump(caminhos, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def caminho_sem_sicoob(caminho_sicoob):
    partes = caminho_sicoob.replace("\\", "/").rstrip("/").split("/")
    if partes and partes[-1].upper() == "SICOOB":
        partes = partes[:-1]
    return "/".join(partes)


def get_time():
    return datetime.datetime.now().strftime("%H:%M:%S")


def renderizar_terminal(placeholder, log_list, nome_sh="conciliacao.sh"):
    html_content = f'''
    <div class="bx-terminal-wrapper">
        <div class="bx-terminal-header">
            <div class="terminal-tab">
                <span class="terminal-dot dot-red"></span>
                <span class="terminal-dot dot-yellow"></span>
                <span class="terminal-dot dot-green"></span>
                <span style="margin-left: 10px;">bash ~ bx_finance/{nome_sh}</span>
            </div>
            <div class="live-badge"><div class="pulse"></div> SERVER LIVE</div>
        </div>
        <div class="bx-terminal-body" id="terminal-body">
    '''

    for linha in log_list:
        if not linha.strip():
            continue

        hora_str = ""
        if linha.startswith("["):
            parts = linha.split("] ", 1)
            if len(parts) == 2:
                hora_str = f'<span class="log-time">{parts[0]}]</span>'
                linha = parts[1]

        linha_safe = html.escape(linha).replace('\n', '<br>')

        css_class = "log-line"
        upper_line = linha.upper()
        if "ERRO" in upper_line or "❌" in linha or "FALHA" in upper_line:
            css_class += " log-error"
        elif "SUCESSO" in upper_line or "🎉" in linha or "✅" in linha or "COMPLETO" in upper_line:
            css_class += " log-success"
        elif "▶️" in linha or "🚀" in linha or "INICIANDO" in upper_line:
            css_class += " log-highlight"
        elif "📁" in linha or "🔐" in linha or "⬆️" in linha or "⬇️" in linha:
            css_class += " log-system"
        elif "Processando:" in linha or "Encontrados" in linha or "Lendo extrato" in linha or "->" in linha:
            css_class += " log-info"

        html_content += f'<div class="{css_class}">{hora_str}{linha_safe}</div>'

    html_content += '''
        </div>
        <script>
            var term = window.parent.document.getElementById('terminal-body');
            if(term) { term.scrollTop = term.scrollHeight; }
        </script>
    </div>
    '''
    placeholder.markdown(html_content, unsafe_allow_html=True)


def rodar_robo_com_log(nome_robo, caminho_script, pasta_trabalho, log_placeholder, log_atual, nome_sh="conciliacao.sh"):
    log_atual.append(f"[{get_time()}] ▶️ [BOOT] INICIANDO MÓDULO: {nome_robo}...")
    renderizar_terminal(log_placeholder, log_atual, nome_sh)

    processo = subprocess.Popen(
        [sys.executable, caminho_script, pasta_trabalho],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        encoding='utf-8',
        errors='replace'
    )

    for linha in processo.stdout:
        clean_line = linha.strip("\n")
        if clean_line:
            log_atual.append(f"[{get_time()}] {clean_line}")
            renderizar_terminal(log_placeholder, log_atual, nome_sh)

    processo.wait()

    if processo.returncode != 0:
        log_atual.append(f"[{get_time()}] ❌ [CRÍTICO] MÓDULO {nome_robo} ENCERROU COM ERRO (CÓDIGO {processo.returncode}).")
        renderizar_terminal(log_placeholder, log_atual, nome_sh)
        return False

    return True


# =====================================================================
# CONTROLE DE SESSÃO (INICIA SEM NENHUM PROJETO SELECIONADO)
# =====================================================================
if "historico_runs" not in st.session_state:
    st.session_state.historico_runs = []
if "mostrar_log_idx" not in st.session_state:
    st.session_state.mostrar_log_idx = None
if "projeto_selecionado" not in st.session_state:
    projeto_url = st.query_params.get("projeto", "").upper()
    if projeto_url in ["CATIVA", "GRUPO CATIVA"]:
        st.session_state.projeto_selecionado = "GRUPO CATIVA"
    elif projeto_url == "REVPACK":
        st.session_state.projeto_selecionado = "REVPACK"
    else:
        # Inicia sem nenhum projeto selecionado (mostra apenas os cards dos projetos)
        st.session_state.projeto_selecionado = None

with st.sidebar:
    st.markdown("<h3 style='color: #6259A4; font-weight: 800; font-family: Inter, sans-serif;'>📜 HISTÓRICO</h3>", unsafe_allow_html=True)
    st.markdown("<hr style='border-color: #1f232b; margin-top: 0;'>", unsafe_allow_html=True)

    if not st.session_state.historico_runs:
        st.markdown("<span style='color: #a0aec0; font-size: 0.85rem;'>Nenhuma execução na sessão atual.</span>", unsafe_allow_html=True)
    else:
        for i, run in enumerate(reversed(st.session_state.historico_runs)):
            idx = len(st.session_state.historico_runs) - 1 - i
            cor_status = "#38A169" if "Sucesso" in run["status"] else "#E53E3E"

            st.markdown(f"**⏰ {run['hora']}** <br> <span style='color: {cor_status}; font-weight: 600; font-size: 0.85rem;'>{run['status']}</span>", unsafe_allow_html=True)
            if st.button(f"🔎 Ver Log", key=f"btn_log_{idx}", use_container_width=True):
                st.session_state.mostrar_log_idx = idx
            st.markdown("<hr style='border-color: #1f232b; margin: 10px 0;'>", unsafe_allow_html=True)


# =====================================================================
# HEADER PRINCIPAL
# =====================================================================
projeto_ativo = st.session_state.projeto_selecionado

if projeto_ativo == "GRUPO CATIVA":
    subtitulo_header = "GRUPO CATIVA — ROBÔ DE CONCILIAÇÃO BANCÁRIA"
    qtd_modulos_card = "03"
elif projeto_ativo == "REVPACK":
    subtitulo_header = "REVPACK — ROBÔ DE CONCILIAÇÃO BANCÁRIA & CONTROLE DE SACADOS"
    qtd_modulos_card = "02"
else:
    subtitulo_header = "PAINEL CENTRAL DE ROBÔS DE CONCILIAÇÃO BANCÁRIA & CONTROLE DE SACADOS"
    qtd_modulos_card = "05"

st.markdown(f'''
<div class="bx-header-container">
    <div class="bx-title-wrapper">
        <div class="bx-central-title">CENTRAL <span class="bx-accent">BX</span> DE AUTOMAÇÕES</div>
        <div class="bx-subtitle-row">
            <span class="bx-subtitle-dash"></span>
            <span class="bx-subtitle-text">{subtitulo_header}</span>
        </div>
    </div>
    <div class="bx-right-logo-wrapper">
        <div class="bx-logo-text"><span class="b">B</span><span class="x">x</span></div>
        <div class="bx-logo-sub">finance</div>
        <div class="bx-build-tag">[ V 2.0 / BUILD SECURE ]</div>
    </div>
</div>
''', unsafe_allow_html=True)

# =====================================================================
# CORPO PRINCIPAL: ESQUERDA (SELECIONE O PROJETO) | DIREITA (CARDS)
# =====================================================================
col_form, col_space, col_metrics = st.columns([5.2, 0.8, 3.0])

caminhos_atuais = carregar_caminhos_salvos()

botao_rodar_cativa = False
botao_rodar_revpack_cb = False
botao_rodar_revpack_cs = False

with col_form:
    st.markdown('<div class="bx-select-title">SELECIONE O PROJETO</div>', unsafe_allow_html=True)

    # Botões estilo Pill: Só fica roxo brilhante (primary) depois que o usuário clicar em um deles
    col_p1, col_p2, _ = st.columns([1.65, 1.45, 2.1])
    with col_p1:
        tipo_btn_cativa = "primary" if projeto_ativo == "GRUPO CATIVA" else "secondary"
        if st.button("GRUPO CATIVA", key="pill_cativa", type=tipo_btn_cativa, use_container_width=True):
            # Se clicar novamente no mesmo, recolhe; se clicar nele fechado, abre o endereçador
            if st.session_state.projeto_selecionado == "GRUPO CATIVA":
                st.session_state.projeto_selecionado = None
            else:
                st.session_state.projeto_selecionado = "GRUPO CATIVA"
            st.rerun()
    with col_p2:
        tipo_btn_revpack = "primary" if projeto_ativo == "REVPACK" else "secondary"
        if st.button("REVPACK", key="pill_revpack", type=tipo_btn_revpack, use_container_width=True):
            if st.session_state.projeto_selecionado == "REVPACK":
                st.session_state.projeto_selecionado = None
            else:
                st.session_state.projeto_selecionado = "REVPACK"
            st.rerun()

    # SÓ EXIBE O ENDEREÇADOR E OS BOTÕES DE RODAR DEPOIS QUE CLICAR EM UM DOS PROJETOS
    if projeto_ativo == "GRUPO CATIVA":
        st.markdown('<div class="bx-module-divider"></div>', unsafe_allow_html=True)
        caminho_digitado_cativa = st.text_input(
            "DIRETÓRIO ALVO NO GOOGLE DRIVE (Referência SICOOB):",
            value=caminhos_atuais["cativa_sicoob"],
            help=r"Exemplo: 2026\10 - Outubro\28.10.2026\SICOOB",
            key="input_cativa_sicoob"
        )
        c_btn, _ = st.columns([2.2, 3.0])
        with c_btn:
            botao_rodar_cativa = st.button("PROCESSAR CONCILIAÇÃO", key="btn_cativa", type="primary", use_container_width=True)

    elif projeto_ativo == "REVPACK":
        st.markdown('<div class="bx-module-divider"></div>', unsafe_allow_html=True)

        # BLOCO 1 (EM CIMA): CONCILIAÇÃO BANCÁRIA REVPACK
        caminho_digitado_rev_cb = st.text_input(
            "DIRETÓRIO ALVO NO GOOGLE DRIVE (CONCILIAÇÃO BANCÁRIA):",
            value=caminhos_atuais["revpack_cb"],
            help=r"Exemplo: EXTRATOS BANCÁRIOS\2026\10. Outubro\28",
            key="input_revpack_cb"
        )
        c_btn_cb, _ = st.columns([2.2, 3.0])
        with c_btn_cb:
            botao_rodar_revpack_cb = st.button("PROCESSAR CONCILIAÇÃO", key="btn_revpack_cb", type="primary", use_container_width=True)

        st.markdown('<div class="bx-module-divider"></div>', unsafe_allow_html=True)

        # BLOCO 2 (EMBAIXO): CONTROLE DE SACADOS REVPACK
        caminho_digitado_rev_cs = st.text_input(
            "DIRETÓRIO ALVO NO GOOGLE DRIVE (CONTROLE DE SACADOS):",
            value=caminhos_atuais["revpack_cs"],
            help=r"Exemplo: POSIÇÕES FUNDOS\2026\10 - OUT\Fundos 29.10\Trabalhados",
            key="input_revpack_cs"
        )
        c_btn_cs, _ = st.columns([2.5, 2.7])
        with c_btn_cs:
            botao_rodar_revpack_cs = st.button("PROCESSAR CONTROLE DE SACADOS", key="btn_revpack_cs", type="primary", use_container_width=True)

with col_metrics:
    st.markdown('<div style="margin-top: 115px;"></div>', unsafe_allow_html=True)
    st.markdown(f'''
    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px;">
        <div class="dash-card">
            <div class="dash-value">{qtd_modulos_card}</div>
            <div class="dash-label">Módulos Auto</div>
        </div>
        <div class="dash-card">
            <div class="dash-value" style="color: #38A169;">OK</div>
            <div class="dash-label">GCP Cloud</div>
        </div>
        <div class="dash-card" style="grid-column: span 2; padding: 22px;">
            <div class="dash-value" style="font-size: 1.35rem; color: #8C82D8; letter-spacing: 2px;">OAUTH 2.0 ATIVO</div>
            <div class="dash-label">Token Security Status</div>
        </div>
    </div>
    ''', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# =====================================================================
# 1. EXECUÇÃO: GRUPO CATIVA (SICOOB + OUTROS)
# =====================================================================
if botao_rodar_cativa:
    st.markdown("<h3 style='color: #ffffff; font-weight: 800; font-size: 1.2rem; margin-bottom: 0px;'>🖥️ TERMINAL DE PROCESSAMENTO</h3>", unsafe_allow_html=True)
    log_placeholder = st.empty()
    log_atual = []
    nome_sh = "cativa_conciliacao.sh"

    renderizar_terminal(log_placeholder, log_atual, nome_sh)
    pasta_temporaria = tempfile.mkdtemp(prefix="bx_cativa_")

    try:
        log_atual.append(f"[{get_time()}] 🔐 [AUTH] Inicializando túnel seguro com Google Drive OAuth 2.0...")
        renderizar_terminal(log_placeholder, log_atual, nome_sh)

        credenciais = json.loads(st.secrets["GOOGLE_TOKEN_JSON"])
        servico = drive_utils.conectar_drive(credenciais)

        log_atual.append(f"[{get_time()}] ✅ [AUTH] Autenticação validada com sucesso.")
        renderizar_terminal(log_placeholder, log_atual, nome_sh)

        log_atual.append(f"[{get_time()}] 📁 [I/O] Varrendo diretório alvo: {caminho_digitado_cativa}")
        renderizar_terminal(log_placeholder, log_atual, nome_sh)

        id_pasta_sicoob = drive_utils.navegar_caminho(
            servico, caminho_digitado_cativa, id_raiz=drive_utils.ID_PASTA_RAIZ_CATIVA
        )
        caminho_outros = caminho_sem_sicoob(caminho_digitado_cativa)
        id_pasta_outros = drive_utils.navegar_caminho(
            servico, caminho_outros, id_raiz=drive_utils.ID_PASTA_RAIZ_CATIVA
        )

        destino_sicoob = os.path.join(pasta_temporaria, "SICOOB")
        destino_outros = pasta_temporaria

        log_atual.append(f"[{get_time()}] ⬇️ [I/O] Efetuando download de matriz SICOOB...")
        renderizar_terminal(log_placeholder, log_atual, nome_sh)
        arquivos_sicoob_originais = drive_utils.baixar_pasta_inteira(servico, id_pasta_sicoob, destino_sicoob)
        log_atual.append(f"[{get_time()}] ✅ [I/O] Matriz SICOOB sincronizada: {len(arquivos_sicoob_originais)} arquivo(s).")

        log_atual.append(f"[{get_time()}] ⬇️ [I/O] Efetuando download de matriz OUTROS BANCOS...")
        renderizar_terminal(log_placeholder, log_atual, nome_sh)
        arquivos_outros_originais = drive_utils.baixar_pasta_inteira(servico, id_pasta_outros, destino_outros)
        log_atual.append(f"[{get_time()}] ✅ [I/O] Matriz OUTROS sincronizada: {len(arquivos_outros_originais)} arquivo(s).")

        log_atual.append("=========================================================")
        log_atual.append("🤖 INICIANDO O ROBÔ DE CONCILIAÇÃO BANCÁRIA (GRUPO CATIVA)")
        log_atual.append("=========================================================")
        log_atual.append(f"📁 Pasta SICOOB configurada: {caminho_digitado_cativa}")
        log_atual.append(f"📁 Pasta OUTROS configurada: {caminho_outros}")
        log_atual.append("-" * 50)
        log_atual.append("🔍 VERIFICANDO ARQUIVOS NA PASTA...")
        renderizar_terminal(log_placeholder, log_atual, nome_sh)

        arquivos_baixados = os.listdir(destino_sicoob)
        divergentes = []
        for arq in arquivos_baixados:
            if arq.startswith("~$"):
                continue
            if "  " in arq:
                divergentes.append(arq)

        if divergentes:
            log_atual.append("")
            log_atual.append("❌ ERRO: Problema na validação dos arquivos antes de iniciar os robôs!")
            log_atual.append("")
            log_atual.append("🟡 DIVERGÊNCIA ENCONTRADA (Nomes incorretos ou arquivos não mapeados):")
            for d in divergentes:
                log_atual.append(f"   - {d}")
            log_atual.append("")
            log_atual.append("➡️ Por favor, ajuste a pasta e rode novamente.")
            renderizar_terminal(log_placeholder, log_atual, nome_sh)

            st.session_state.historico_runs.append({"hora": get_time(), "status": "❌ Divergência (Grupo Cativa)", "log": list(log_atual)})
            st.session_state.mostrar_log_idx = len(st.session_state.historico_runs) - 1
            st.error("O processo foi interrompido. Verifique os nomes incorretos no terminal, corrija os espaços duplos no Google Drive e tente novamente.")
            st.stop()

        log_atual.append("✅ Todos os arquivos obrigatórios estão presentes e não há nomes divergentes!")
        log_atual.append("-" * 50)
        renderizar_terminal(log_placeholder, log_atual, nome_sh)

        ok = rodar_robo_com_log("ROBÔ 1 (PDF + PIX)", SCRIPT_CATIVA_PDF, destino_sicoob, log_placeholder, log_atual, nome_sh)

        if ok:
            ok = rodar_robo_com_log("ROBÔ 2 (CONCILIAÇÃO SICOOB)", SCRIPT_CATIVA_SICOOB, destino_sicoob, log_placeholder, log_atual, nome_sh)

        if ok:
            ok = rodar_robo_com_log("ROBÔ OUTROS BANCOS", SCRIPT_CATIVA_OUTROS, destino_outros, log_placeholder, log_atual, nome_sh)

        if not ok:
            st.session_state.historico_runs.append({"hora": get_time(), "status": "❌ Falha (Grupo Cativa)", "log": list(log_atual)})
        else:
            log_atual.append(f"[{get_time()}] ⬆️ [I/O] Transmitindo relatórios finais (SICOOB) para a nuvem...")
            renderizar_terminal(log_placeholder, log_atual, nome_sh)
            drive_utils.subir_arquivos_novos(servico, id_pasta_sicoob, destino_sicoob, arquivos_sicoob_originais)

            log_atual.append(f"[{get_time()}] ⬆️ [I/O] Transmitindo relatórios finais (OUTROS) para a nuvem...")
            renderizar_terminal(log_placeholder, log_atual, nome_sh)
            drive_utils.subir_arquivos_novos(servico, id_pasta_outros, destino_outros, arquivos_outros_originais)

            log_atual.append("")
            log_atual.append("=========================================================")
            log_atual.append("✅ PROCESSO CONCLUÍDO COM SUCESSO! - Os relatórios gerados já estão disponíveis no seu Google Drive.")
            log_atual.append("=========================================================")
            renderizar_terminal(log_placeholder, log_atual, nome_sh)

            salvar_caminho_especifico("cativa_sicoob", caminho_digitado_cativa)
            st.session_state.historico_runs.append({"hora": get_time(), "status": "✅ Sucesso (Grupo Cativa)", "log": list(log_atual)})

    except FileNotFoundError as e:
        log_atual.append(f"[{get_time()}] ❌ [ERRO DE DIRETÓRIO] Pasta não localizada no Drive: {e}")
        renderizar_terminal(log_placeholder, log_atual, nome_sh)
        st.session_state.historico_runs.append({"hora": get_time(), "status": "❌ Erro de Pasta (Grupo Cativa)", "log": list(log_atual)})
        st.session_state.mostrar_log_idx = len(st.session_state.historico_runs) - 1
        st.error(f"Não encontrei uma das pastas no Drive: {e}")
    except Exception as e:
        log_atual.append(f"[{get_time()}] ❌ [ERRO DE SISTEMA] Falha inesperada: {e}")
        renderizar_terminal(log_placeholder, log_atual, nome_sh)
        st.session_state.historico_runs.append({"hora": get_time(), "status": "❌ Erro de Sistema (Grupo Cativa)", "log": list(log_atual)})
        st.session_state.mostrar_log_idx = len(st.session_state.historico_runs) - 1
        st.error(f"Ocorreu um erro inesperado: {e}")
    finally:
        shutil.rmtree(pasta_temporaria, ignore_errors=True)


# =====================================================================
# 2. EXECUÇÃO: REVPACK (CONCILIAÇÃO BANCÁRIA OU CONTROLE DE SACADOS)
# =====================================================================
elif botao_rodar_revpack_cb or botao_rodar_revpack_cs:
    if botao_rodar_revpack_cb:
        nome_operacao = "Conciliação Bancária"
        nome_modulo = "AUT - CB - REVPACK"
        script_alvo = SCRIPT_REVPACK_CB
        caminho_alvo = caminho_digitado_rev_cb
        id_raiz_alvo = drive_utils.ID_PASTA_REVPACK_CONCILIACAO
        chave_salvar = "revpack_cb"
        nome_sh = "revpack_conciliacao.sh"
    else:
        nome_operacao = "Controle de Sacados"
        nome_modulo = "CONTROLE DE SACADOS - REVPACK"
        script_alvo = SCRIPT_REVPACK_CS
        caminho_alvo = caminho_digitado_rev_cs
        id_raiz_alvo = drive_utils.ID_PASTA_REVPACK_SACADOS
        chave_salvar = "revpack_cs"
        nome_sh = "revpack_sacados.sh"

    st.markdown("<h3 style='color: #ffffff; font-weight: 800; font-size: 1.2rem; margin-bottom: 0px;'>🖥️ TERMINAL DE PROCESSAMENTO</h3>", unsafe_allow_html=True)
    log_placeholder = st.empty()
    log_atual = []

    renderizar_terminal(log_placeholder, log_atual, nome_sh)
    pasta_temporaria = tempfile.mkdtemp(prefix="bx_revpack_")

    try:
        log_atual.append(f"[{get_time()}] 🔐 [AUTH] Inicializando túnel seguro com Google Drive OAuth 2.0...")
        renderizar_terminal(log_placeholder, log_atual, nome_sh)

        credenciais = json.loads(st.secrets["GOOGLE_TOKEN_JSON"])
        servico = drive_utils.conectar_drive(credenciais)

        log_atual.append(f"[{get_time()}] ✅ [AUTH] Autenticação validada com sucesso.")
        renderizar_terminal(log_placeholder, log_atual, nome_sh)

        log_atual.append(f"[{get_time()}] 📁 [I/O] Varrendo diretório alvo ({nome_operacao}): {caminho_alvo}")
        renderizar_terminal(log_placeholder, log_atual, nome_sh)

        id_pasta_final = drive_utils.navegar_caminho(servico, caminho_alvo, id_raiz=id_raiz_alvo)

        log_atual.append(f"[{get_time()}] ⬇️ [I/O] Efetuando download dos arquivos da pasta alvo...")
        renderizar_terminal(log_placeholder, log_atual, nome_sh)
        arquivos_originais = drive_utils.baixar_pasta_inteira(servico, id_pasta_final, pasta_temporaria)
        log_atual.append(f"[{get_time()}] ✅ [I/O] Pasta sincronizada: {len(arquivos_originais)} arquivo(s) baixado(s).")

        log_atual.append("=========================================================")
        log_atual.append(f"🤖 INICIANDO {nome_operacao.upper()} (REVPACK)")
        log_atual.append("=========================================================")
        log_atual.append(f"📁 Diretório configurado: {caminho_alvo}")
        log_atual.append("-" * 50)
        renderizar_terminal(log_placeholder, log_atual, nome_sh)

        ok = rodar_robo_com_log(nome_modulo, script_alvo, pasta_temporaria, log_placeholder, log_atual, nome_sh)

        if not ok:
            st.session_state.historico_runs.append({"hora": get_time(), "status": f"❌ Falha (Revpack {nome_operacao})", "log": list(log_atual)})
            st.session_state.mostrar_log_idx = len(st.session_state.historico_runs) - 1
        else:
            log_atual.append(f"[{get_time()}] ⬆️ [I/O] Transmitindo relatórios finais para a nuvem...")
            renderizar_terminal(log_placeholder, log_atual, nome_sh)
            enviados = drive_utils.subir_arquivos_novos(
                servico, id_pasta_final, pasta_temporaria, arquivos_originais, apenas_novos=True
            )
            log_atual.append(f"[{get_time()}] ✅ [I/O] Relatórios sincronizados no Drive: {', '.join(enviados) if enviados else 'Concluído'}")

            log_atual.append("")
            log_atual.append("=========================================================")
            log_atual.append("✅ PROCESSO CONCLUÍDO COM SUCESSO! - Os relatórios gerados já estão disponíveis no seu Google Drive.")
            log_atual.append("=========================================================")
            renderizar_terminal(log_placeholder, log_atual, nome_sh)

            salvar_caminho_especifico(chave_salvar, caminho_alvo)
            st.session_state.historico_runs.append({"hora": get_time(), "status": f"✅ Sucesso (Revpack {nome_operacao})", "log": list(log_atual)})

    except FileNotFoundError as e:
        log_atual.append(f"[{get_time()}] ❌ [ERRO DE DIRETÓRIO] Pasta não localizada no Drive: {e}")
        renderizar_terminal(log_placeholder, log_atual, nome_sh)
        st.session_state.historico_runs.append({"hora": get_time(), "status": "❌ Erro de Pasta (Revpack)", "log": list(log_atual)})
        st.session_state.mostrar_log_idx = len(st.session_state.historico_runs) - 1
        st.error(f"Não encontrei uma das pastas no Drive: {e}")
    except Exception as e:
        log_atual.append(f"[{get_time()}] ❌ [ERRO DE SISTEMA] Falha inesperada: {e}")
        renderizar_terminal(log_placeholder, log_atual, nome_sh)
        st.session_state.historico_runs.append({"hora": get_time(), "status": "❌ Erro de Sistema (Revpack)", "log": list(log_atual)})
        st.session_state.mostrar_log_idx = len(st.session_state.historico_runs) - 1
        st.error(f"Ocorreu um erro inesperado: {e}")
    finally:
        shutil.rmtree(pasta_temporaria, ignore_errors=True)


# =====================================================================
# RENDERIZAR LOG HISTÓRICO SE SELECIONADO NA LATERAL
# =====================================================================
elif st.session_state.mostrar_log_idx is not None:
    st.markdown("<h3 style='color: #ffffff; font-weight: 800; font-size: 1.2rem; margin-bottom: 0px;'>🖥️ TERMINAL DE PROCESSAMENTO (HISTÓRICO)</h3>", unsafe_allow_html=True)
    log_placeholder = st.empty()
    idx = st.session_state.mostrar_log_idx
    log_salvo = st.session_state.historico_runs[idx]["log"]
    renderizar_terminal(log_placeholder, log_salvo)
