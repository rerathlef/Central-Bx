import os
import re
import sys
import warnings
import unicodedata
from datetime import datetime
from html.parser import HTMLParser
import pandas as pd

warnings.simplefilter("ignore")

print("--- 🚀 INICIANDO ROBÔ: AUT - CB - REVPACK (V3) ---")

# =============================================================================
# 1. CONFIGURAÇÕES E ENDEREÇO DA PASTA (ALTERE AQUI)
# =============================================================================
if len(sys.argv) > 1 and not sys.argv[1].startswith("-"):
    PASTA_ARQUIVOS = sys.argv[1]
else:
    # Coloque abaixo o caminho da pasta onde estão os extratos da Revpack:
    PASTA_ARQUIVOS = r"G:\.shortcut-targets-by-id\1HH8YPWOEhDI7x9lk-LNQXlVf9sUusDFg\08 Revpack\03 Financeiro\02. EXTRATOS\2026\10. Outubro\26\conciliado\TESTE1"

NOME_RELATORIO_FINAL = "AUT - CONCILIAÇÃO - REVPACK.xlsx"

# Se quiser filtrar apenas um dia específico (ex: "26/10/2026"), preencha abaixo.
# Deixe "" (vazio) para extrair todos os dias presentes nos extratos:
FILTRAR_DATA = ""

# False = Insere a linha "SALDO FINAL" abaixo das movimentações de cada dia (padrão Cativa V8)
# True  = Insere apenas 1 linha "SALDO FINAL" no final de cada extrato
APENAS_ULTIMO_SALDO_POR_EXTRATO = False

# Lista de termos de saldo originais do banco para limpar antes de gerar o nosso SALDO FINAL
TERMOS_EXCLUIR = [
    "S A L D O", "SALDO ANTERIOR", "SALDO INICIAL", "SALDO FINAL",
    "S.A.L.D.O", "SALDO DO DIA", "SALDO INVEST FACIL", "SALDO INVEST FÁCIL"
]

# =============================================================================
# 2. CONFIGURAÇÃO DOS BANCOS DA REVPACK
# =============================================================================
# regra_entrada:
#   - "VAZIO": Deixa o FAVORECIDO em branco nas entradas (Valor > 0) -> Invista e Bradesco
#   - "REVPACK": Preenche "REVPACK" no FAVORECIDO em todas as entradas (Valor > 0) -> Daycoval e Continental
CONFIG_BANCOS = {
    "DAYCOVAL": {
        "id_arquivo": "DAYCOVAL",
        "col_data": "Data",
        "col_desc": "Lancamento",
        "col_debito": "Debito (R$)",
        "col_credito": "Credito (R$)",
        "col_saldo": "Saldo (R$)",
        "tipo_valor": "DUPLA",
        "regra_saldo": "ULTIMO_VALIDO_DIA",
        "ordem_extrato": "CRESCENTE",
        "regra_entrada": "REVPACK",
        "nome_banco_padrao": "BANCO DAYCOVAL"
    },
    "BRADESCO": {
        "id_arquivo": "BRADESCO",
        "col_data": "Data",
        "col_desc": "Lançamento",
        "col_credito": "Crédito (R$)",
        "col_debito": "Débito (R$)",
        "col_saldo": "Saldo (R$)",
        "tipo_valor": "DUPLA_SMART",
        "regra_saldo": "ULTIMO_VALIDO_DIA",
        "ordem_extrato": "CRESCENTE",
        "parar_em": ["TOTAL", "SALDOS INVEST"],
        "regra_entrada": "VAZIO",
        "nome_banco_padrao": "BANCO BRADESCO"
    },
    "CONTINENTAL": {
        "id_arquivo": "CONTINENTAL",
        "col_data": "Data_da_Ocorrencia",
        "col_desc": "Lancamento",
        "col_favorecido_raw": "Nome",
        "col_valor_1": "Valor",
        "col_saldo": "Saldo",
        "tipo_valor": "UNICA",
        "regra_saldo": "EXTRAIR_TOPO",
        "ordem_extrato": "DECRESCENTE",
        "regra_entrada": "REVPACK",
        "nome_banco_padrao": "BANCO CONTINENTAL"
    },
    "INVISTA": {
        "id_arquivo": "INVISTA",
        "col_data": "Data",
        "col_tipo": "Tipo",
        "col_desc": "Descrição",
        "col_valor_1": "Valor R$",
        "tipo_valor": "UNICA",
        "regra_saldo": "LINHA_SALDO_DO_DIA",
        "ordem_extrato": "DECRESCENTE",
        "regra_entrada": "VAZIO",
        "nome_banco_padrao": "BANCO INVISTA"
    }
}

# =============================================================================
# 3. REGRAS AVANÇADAS DE FAVORECIDOS (APLICADAS NAS SAÍDAS / VALOR < 0)
# =============================================================================
REGRAS_AVANCADAS_SAIDA = [
    # Tarifas e Encargos -> Nome do próprio Banco
    {"termo": "TARIFAS DE CONTA", "tipo": "FIXO_BANCO"},
    {"termo": "TARIFA DE TED", "tipo": "FIXO_BANCO"},
    {"termo": "TAR COB TITULOS", "tipo": "FIXO_BANCO"},
    {"termo": "TAR. DOC / TED", "tipo": "FIXO_BANCO"},
    {"termo": "TAR BLOQ/DESBLOQ", "tipo": "FIXO_BANCO"},
    {"termo": "TARIFA", "tipo": "FIXO_BANCO"},
    {"termo": "TAR ", "tipo": "FIXO_BANCO"},
    {"termo": "TAR.", "tipo": "FIXO_BANCO"},
    {"termo": "IOF", "tipo": "FIXO", "valor": "RECEITA FEDERAL"},

    # Regras específicas de Saída (Daycoval / Bradesco)
    {"termo": "TRANSF.MESMA TITULARIDADE", "tipo": "FIXO", "valor": "REVPACK"},
    {"termo": "BLOQUEIO-ORDEM JUDICIAL", "tipo": "FIXO", "valor": "BLOQUEIO JUDICIAL"},
    {"termo": "CAMARA DE COM DE ENERGIA ELET", "tipo": "FIXO", "valor": "CCEE - CAMARA DE COM DE ENERGIA"},
    {"termo": "TRANSFERENCIA PIX DES:", "tipo": "EXTRAIR_APOS_TERMO"},
]

# =============================================================================
# 4. FUNÇÕES DE APOIO E LEITURA (EXCEL + HTML DAYCOVAL)
# =============================================================================
class LeitorTabelaHTML(HTMLParser):
    """Lê arquivos .xls que internamente estão salvos como HTML (padrão Daycoval)."""
    def __init__(self):
        super().__init__()
        self.in_table = False
        self.in_row = False
        self.in_cell = False
        self.tables = []
        self.curr_table = []
        self.curr_row = []
        self.curr_cell = []

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.in_table = True
            self.curr_table = []
        elif tag == "tr" and self.in_table:
            self.in_row = True
            self.curr_row = []
        elif tag in ("td", "th") and self.in_row:
            self.in_cell = True
            self.curr_cell = []

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self.in_cell:
            self.in_cell = False
            self.curr_row.append("".join(self.curr_cell).strip())
        elif tag == "tr" and self.in_row:
            self.in_row = False
            if any(c != "" for c in self.curr_row):
                self.curr_table.append(self.curr_row)
        elif tag == "table" and self.in_table:
            self.in_table = False
            if self.curr_table:
                self.tables.append(self.curr_table)

    def handle_data(self, data):
        if self.in_cell:
            self.curr_cell.append(data)


def limpar_valor_moeda(valor):
    """Converte textos de moeda para float; retorna None se estiver vazio ou '-'."""
    if valor is None or pd.isna(valor):
        return None
    v = str(valor).strip().replace("R$", "").replace("\xa0", "").replace(" ", "").replace('"', "").replace("'", "")
    if v in ("", "-", "NAN", "NONE", "NULL"):
        return None
    if "," in v and "." in v:
        v = v.replace(".", "").replace(",", ".")
    elif "," in v:
        v = v.replace(",", ".")
    try:
        return round(float(v), 2)
    except:
        return None


def date_to_excel_serial(date_val):
    """Converte datetime para o número serial do Excel (base 30/12/1899)."""
    if pd.isna(date_val):
        return ""
    try:
        dt_limpa = pd.Timestamp(date_val).to_pydatetime().replace(tzinfo=None)
        return (dt_limpa - datetime(1899, 12, 30)).days
    except:
        return ""


def parse_data_flex(val):
    """Converte diferentes formatos de data (DD/MM/YYYY, DD-MM-YYYY, YYYY-MM-DD HH:MM:SS)."""
    if pd.isna(val):
        return pd.NaT
    s = str(val).strip()
    if s.upper() in ("", "NAN", "NONE", "NAT"):
        return pd.NaT
    s_clean = s.split(" ")[0].replace("-", "/")
    if re.match(r"^\d{4}/\d{2}/\d{2}$", s_clean):
        return pd.to_datetime(s_clean, format="%Y/%m/%d", errors="coerce")
    return pd.to_datetime(s_clean, dayfirst=True, errors="coerce")


def normalizar_texto(texto):
    if not isinstance(texto, str):
        return str(texto)
    return "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn").upper().strip()


def encontrar_coluna_smart(df, palavra_chave):
    if not palavra_chave:
        return None
    alvo = normalizar_texto(palavra_chave)
    for col in df.columns:
        if alvo in normalizar_texto(col):
            return col
    return None


def encontrar_cabecalho(caminho, nome_coluna_data):
    try:
        df_temp = pd.read_excel(caminho, nrows=40, header=None)
        alvo = normalizar_texto(nome_coluna_data)
        for i, row in df_temp.iterrows():
            linha_texto = [normalizar_texto(x) for x in row.values]
            if alvo in linha_texto:
                return i
    except:
        pass
    return None


def carregar_dataframe_extrato(caminho, config):
    """Detecta automaticamente se o arquivo é HTML (.xls do Daycoval) ou Excel padrão."""
    with open(caminho, "rb") as fp:
        inicio = fp.read(200).lstrip()

    if inicio.startswith(b"<") or b"<html" in inicio.lower() or b"<!doctype html" in inicio.lower():
        with open(caminho, "r", encoding="utf-8", errors="replace") as fp:
            conteudo = fp.read()

        # Captura o ano na linha "Gerado em: DD/MM/YYYY" do Daycoval
        m_ano = re.search(r"Gerado em:.*?(\d{2}/\d{2}/(\d{4}))", conteudo, re.IGNORECASE | re.DOTALL)
        if not m_ano:
            m_ano = re.search(r"(\d{2}/\d{2}/(\d{4}))", conteudo)
        ano_ref = m_ano.group(2) if m_ano else str(datetime.now().year)

        parser = LeitorTabelaHTML()
        parser.feed(conteudo)
        alvo_col = normalizar_texto(config["col_data"])
        for t in parser.tables:
            if len(t) > 1 and any(alvo_col == normalizar_texto(c) for c in t[0]):
                headers = [c.strip() for c in t[0]]
                df = pd.DataFrame(t[1:], columns=headers)
                col_d = encontrar_coluna_smart(df, config["col_data"])
                if col_d:
                    df[col_d] = df[col_d].apply(
                        lambda x: f"{str(x).strip()}/{ano_ref}" if re.match(r"^\d{2}/\d{2}$", str(x).strip()) else x
                    )
                return df
        return pd.DataFrame()
    else:
        cabecalho = encontrar_cabecalho(caminho, config["col_data"])
        if cabecalho is None:
            return pd.DataFrame()
        df = pd.read_excel(caminho, skiprows=cabecalho)
        df.columns = [str(c).strip() for c in df.columns]
        return df


def limpar_nome_empresa(texto):
    """Remove códigos bancários, agência, conta e CNPJ/CPF, mantendo apenas o nome da empresa."""
    if pd.isna(texto) or str(texto).strip().upper() in ("", "NAN", "NONE"):
        return ""
    t = str(texto).strip()
    # Se houver " - ", pega o trecho final após o traço (onde fica o nome da empresa)
    if " - " in t:
        t = t.split(" - ")[-1].strip()
    # Remove padrões de CNPJ (00.000.000/0000-00) e CPF (000.000.000-00)
    t = re.sub(r"\b\d{2,3}\.?\d{3}\.?\d{3}[/-]\d{4}-?\d{2}\b", "", t)
    t = re.sub(r"\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b", "", t)
    # Remove números de banco/agência/conta no início (ex: "00019 0007182068 ")
    t = re.sub(r"^[\d\s\./-]+", "", t).strip()
    t = re.sub(r"\s+", " ", t).strip(" -")
    return t.upper()


def extrair_favorecido(lancamento, fav_raw, valor, config):
    # -------------------------------------------------------------------------
    # 1. REGRA PARA ENTRADAS (Valor > 0)
    # -------------------------------------------------------------------------
    if valor > 0:
        regra_ent = config.get("regra_entrada", "VAZIO")
        if regra_ent == "REVPACK":
            return "REVPACK"
        return ""

    # -------------------------------------------------------------------------
    # 2. REGRA PARA SAÍDAS (Valor < 0)
    # -------------------------------------------------------------------------
    desc_upper = normalizar_texto(lancamento)
    fav_upper = normalizar_texto(fav_raw) if pd.notna(fav_raw) else ""

    # Verifica primeiro se bate em alguma Regra Avançada (Tarifas, IOF, CCEE, Bloqueio, etc.)
    for regra in REGRAS_AVANCADAS_SAIDA:
        termo_norm = normalizar_texto(regra["termo"])
        if termo_norm in desc_upper or termo_norm in fav_upper:
            if regra["tipo"] == "FIXO_BANCO":
                return config["nome_banco_padrao"]
            elif regra["tipo"] == "FIXO":
                return regra["valor"]
            elif regra["tipo"] == "EXTRAIR_APOS_TERMO":
                idx = desc_upper.find(termo_norm)
                if idx != -1:
                    resto = str(lancamento)[idx + len(regra["termo"]):].strip()
                    return limpar_nome_empresa(resto)

    # INVISTA: Retorna o nome da empresa já limpo (sem CNPJ/CPF/agência/conta)
    if config["id_arquivo"] == "INVISTA":
        return limpar_nome_empresa(fav_raw)

    # CONTINENTAL: Traz direto da coluna "Nome" do extrato
    if config["id_arquivo"] == "CONTINENTAL":
        if pd.notna(fav_raw) and str(fav_raw).strip().upper() not in ("", "NAN", "NONE"):
            return str(fav_raw).strip().upper()
        return ""

    # DAYCOVAL: Extrai o favorecido que aparece na coluna "Lancamento" do extrato
    if config["id_arquivo"] == "DAYCOVAL":
        nome_limpo = limpar_nome_empresa(lancamento)
        return nome_limpo if nome_limpo else str(lancamento).strip().upper()

    # BRADESCO e demais:
    nome_limpo = limpar_nome_empresa(fav_raw)
    return nome_limpo if nome_limpo else str(lancamento).strip().upper()


# =============================================================================
# 5. MOTOR PRINCIPAL DE PROCESSAMENTO
# =============================================================================
def processar_revpack():
    if not os.path.exists(PASTA_ARQUIVOS):
        print(f"❌ Pasta não encontrada: {PASTA_ARQUIVOS}")
        return

    arquivos = sorted([
        f for f in os.listdir(PASTA_ARQUIVOS)
        if f.lower().endswith((".xlsx", ".xls"))
        and not f.startswith("~$")
        and f != NOME_RELATORIO_FINAL
    ])

    if not arquivos:
        print("⚠️ Nenhum arquivo Excel (.xls / .xlsx) encontrado na pasta configurada.")
        return

    dados_consolidados = []
    abas_individuais = {}
    relatorio_status = {}

    for arquivo in arquivos:
        origem = os.path.splitext(arquivo)[0].strip()
        config = next((cfg for _, cfg in CONFIG_BANCOS.items() if cfg["id_arquivo"] in origem.upper()), None)

        if not config:
            relatorio_status[arquivo] = {"STATUS": "NÃO MAPEADO", "LINHAS": 0, "OBS": "Banco não reconhecido no nome"}
            continue

        try:
            print(f"➜ Processando: {arquivo}")
            caminho = os.path.join(PASTA_ARQUIVOS, arquivo)
            df_orig = carregar_dataframe_extrato(caminho, config)

            if df_orig.empty:
                relatorio_status[arquivo] = {"STATUS": "VAZIO/ERRO", "LINHAS": 0, "OBS": "Cabeçalho não encontrado"}
                continue

            # 1. Trava de rodapé (Ex: Bradesco tem segunda tabela de Invest Fácil abaixo do Total)
            if "parar_em" in config:
                idx_parada = None
                for idx, row in df_orig.iterrows():
                    texto_linha = " ".join(str(v).upper() for v in row.values if pd.notna(v))
                    if any(termo in texto_linha for termo in config["parar_em"]):
                        idx_parada = idx
                        break
                if idx_parada is not None:
                    df_orig = df_orig.loc[:idx_parada - 1].copy()

            # 2. Captura de saldos explícitos em linhas dedicadas (Padrão Invista: "Saldo do dia DD-MM-YYYY")
            saldos_invista = {}
            if config["regra_saldo"] == "LINHA_SALDO_DO_DIA":
                col_desc_inv = encontrar_coluna_smart(df_orig, config["col_desc"])
                col_saldo_inv = df_orig.columns[4] if len(df_orig.columns) > 4 else None
                if col_desc_inv and col_saldo_inv:
                    for _, row in df_orig.iterrows():
                        txt = str(row.get(col_desc_inv, "")).strip()
                        m_sd = re.search(r"Saldo do dia\s*(\d{2}[-/]\d{2}[-/]\d{4})", txt, re.IGNORECASE)
                        if m_sd:
                            dt_sd = parse_data_flex(m_sd.group(1))
                            val_sd = limpar_valor_moeda(row.get(col_saldo_inv))
                            if pd.notna(dt_sd) and val_sd is not None:
                                saldos_invista[date_to_excel_serial(dt_sd)] = val_sd

            # 3. Mapeamento das colunas e datas
            col_data = encontrar_coluna_smart(df_orig, config["col_data"])
            col_desc = encontrar_coluna_smart(df_orig, config["col_desc"])
            col_saldo = encontrar_coluna_smart(df_orig, config.get("col_saldo", ""))

            df_orig["DATA_OBJ"] = df_orig[col_data].apply(parse_data_flex)
            df_orig["DATA_SERIAL"] = df_orig["DATA_OBJ"].apply(date_to_excel_serial)

            df_proc = df_orig.dropna(subset=["DATA_OBJ"]).copy()

            # 4. Construção do Lançamento Extrato e Favorecido bruto por banco
            if config["id_arquivo"] == "INVISTA":
                # INVISTA: "Tipo" + " - " + "Descrição limpa (sem CNPJ/CPF/agência/conta)"
                col_tipo = encontrar_coluna_smart(df_proc, config["col_tipo"])
                tipos = df_proc[col_tipo].astype(str).fillna("").str.strip()
                descs_raw = df_proc[col_desc].apply(
                    lambda x: "" if pd.isna(x) or str(x).strip().upper() == "NAN" else str(x).strip()
                )
                descs_limpas = [limpar_nome_empresa(d) for d in descs_raw]
                df_proc["LANC_FINAL"] = [f"{t} - {d}" if d else t for t, d in zip(tipos, descs_limpas)]
                df_proc["FAV_RAW"] = descs_limpas

            elif config["id_arquivo"] == "CONTINENTAL":
                # CONTINENTAL: "Lancamento" + " - " + "Nome"
                col_fav = encontrar_coluna_smart(df_proc, config.get("col_favorecido_raw", ""))
                lancs = df_proc[col_desc].astype(str).fillna("").str.strip()
                nomes = df_proc[col_fav].apply(
                    lambda x: "" if pd.isna(x) or str(x).strip().upper() in ("", "NAN", "NONE") else str(x).strip()
                ) if col_fav else [""] * len(df_proc)
                df_proc["LANC_FINAL"] = [f"{l} - {n}" if n else l for l, n in zip(lancs, nomes)]
                df_proc["FAV_RAW"] = nomes

            else:
                df_proc["LANC_FINAL"] = df_proc[col_desc].astype(str).str.strip()
                col_fav = encontrar_coluna_smart(df_proc, config.get("col_favorecido_raw", ""))
                df_proc["FAV_RAW"] = df_proc[col_fav] if col_fav else df_proc["LANC_FINAL"]

            # Remove linhas de saldo originais do banco para não duplicar
            padrao_excl = "|".join(TERMOS_EXCLUIR)
            df_proc = df_proc[~df_proc["LANC_FINAL"].str.upper().str.contains(padrao_excl, na=False)].copy()

            # 5. Cálculo dos Valores (Coluna Única ou Crédito/Débito separados)
            if config["tipo_valor"] in ("DUPLA", "DUPLA_SMART"):
                col_c = encontrar_coluna_smart(df_proc, config["col_credito"])
                col_d = encontrar_coluna_smart(df_proc, config["col_debito"])
                vals = []
                for _, r in df_proc.iterrows():
                    vc = limpar_valor_moeda(r.get(col_c)) if col_c else None
                    vd = limpar_valor_moeda(r.get(col_d)) if col_d else None
                    if vd is not None and vd != 0:
                        vals.append(-abs(vd))
                    elif vc is not None:
                        vals.append(abs(vc))
                    else:
                        vals.append(0.0)
                df_proc["VALOR_FINAL"] = vals
            else:
                col_v = encontrar_coluna_smart(df_proc, config["col_valor_1"])
                df_proc["VALOR_FINAL"] = df_proc[col_v].apply(lambda x: limpar_valor_moeda(x) or 0.0)

            df_proc = df_proc[df_proc["VALOR_FINAL"] != 0].copy()

            # Filtro opcional de data
            if FILTRAR_DATA:
                alvo_serial = date_to_excel_serial(parse_data_flex(FILTRAR_DATA))
                df_proc = df_proc[df_proc["DATA_SERIAL"] == alvo_serial].copy()

            if df_proc.empty:
                relatorio_status[arquivo] = {"STATUS": "SEM DADOS", "LINHAS": 0, "OBS": "Sem movimentação válida"}
                continue

            # 6. Aplicação das regras de Favorecido (Entradas x Saídas por Banco)
            df_proc["FAV_FINAL"] = [
                extrair_favorecido(r["LANC_FINAL"], r["FAV_RAW"], r["VALOR_FINAL"], config)
                for _, r in df_proc.iterrows()
            ]

            ordem_mult = -1 if config.get("ordem_extrato") == "DECRESCENTE" else 1

            df_aba = pd.DataFrame({
                "AG./ORIGEM": origem,
                " ": "",
                "DATA_SERIAL": df_proc["DATA_SERIAL"],
                "DATA": df_proc["DATA_SERIAL"],
                "LANÇAMENTO EXTRATO": df_proc["LANC_FINAL"],
                "VALOR (R$)": df_proc["VALOR_FINAL"],
                "  ": "",
                "FAVORECIDO": df_proc["FAV_FINAL"],
                "CHECK": "",
                "priority": 0,
                "original_index": df_proc.index,
                "sort_order": df_proc.index * ordem_mult
            })

            # 7. Extração do SALDO FINAL do dia (caso exista no extrato)
            saldos_fechamento = []
            datas_unicas = sorted(df_aba["DATA_SERIAL"].unique())
            if APENAS_ULTIMO_SALDO_POR_EXTRATO and datas_unicas:
                datas_unicas = [datas_unicas[-1]]

            for data_serial in datas_unicas:
                subset_dia = df_aba[df_aba["DATA_SERIAL"] == data_serial]
                valor_saldo = None

                if config["regra_saldo"] == "EXTRAIR_TOPO" and col_saldo:
                    idx_topo = subset_dia["original_index"].min()
                    valor_saldo = limpar_valor_moeda(df_orig.loc[idx_topo, col_saldo])

                elif config["regra_saldo"] == "ULTIMO_VALIDO_DIA" and col_saldo:
                    indices_dia = sorted(subset_dia["original_index"].tolist())
                    saldos_validos = [
                        limpar_valor_moeda(df_orig.loc[i, col_saldo])
                        for i in indices_dia
                        if limpar_valor_moeda(df_orig.loc[i, col_saldo]) is not None
                    ]
                    if saldos_validos:
                        valor_saldo = saldos_validos[-1]

                elif config["regra_saldo"] == "LINHA_SALDO_DO_DIA":
                    valor_saldo = saldos_invista.get(data_serial, None)

                if valor_saldo is not None:
                    saldos_fechamento.append({
                        "AG./ORIGEM": origem,
                        " ": "",
                        "DATA_SERIAL": data_serial,
                        "DATA": data_serial,
                        "LANÇAMENTO EXTRATO": "SALDO FINAL",
                        "VALOR (R$)": valor_saldo,
                        "  ": "",
                        "FAVORECIDO": "",
                        "CHECK": "",
                        "priority": 1,
                        "original_index": 999999,
                        "sort_order": 999999
                    })

            if saldos_fechamento:
                df_aba = pd.concat([df_aba, pd.DataFrame(saldos_fechamento)], ignore_index=True)

            df_aba = (
                df_aba.sort_values(by=["DATA_SERIAL", "priority", "sort_order"])
                .drop(columns=["priority", "original_index", "sort_order", "DATA_SERIAL"])
                .fillna("")
            )

            dados_consolidados.append(df_aba)
            abas_individuais[origem[:30]] = df_aba
            relatorio_status[arquivo] = {"STATUS": "SUCESSO", "LINHAS": len(df_aba), "OBS": "OK"}

        except Exception as e:
            relatorio_status[arquivo] = {"STATUS": "ERRO", "LINHAS": 0, "OBS": str(e)[:45]}

    # =========================================================================
    # 6. RESUMO NO CONSOLE E EXPORTAÇÃO DO EXCEL
    # =========================================================================
    print("\n" + "=" * 95)
    print(f"{'RESUMO DO PROCESSAMENTO - REVPACK':^95}")
    print("=" * 95)
    print(f"{'ARQUIVO':<42} | {'STATUS':<12} | {'LINHAS':<6} | {'OBSERVAÇÃO'}")
    print("-" * 95)
    for arq, info in relatorio_status.items():
        print(f"{arq[:40]:<42} | {info['STATUS']:<12} | {info['LINHAS']:<6} | {info['OBS']}")
    print("=" * 95)

    if dados_consolidados:
        df_completo = pd.concat(dados_consolidados, ignore_index=True)
        caminho_final = os.path.join(PASTA_ARQUIVOS, NOME_RELATORIO_FINAL)

        with pd.ExcelWriter(caminho_final, engine="openpyxl") as writer:
            df_completo.to_excel(writer, sheet_name="CONSOLIDADO", index=False)
            ws_cons = writer.sheets["CONSOLIDADO"]
            ws_cons.column_dimensions["A"].width = 32
            ws_cons.column_dimensions["C"].width = 12
            ws_cons.column_dimensions["D"].width = 68
            ws_cons.column_dimensions["E"].width = 16
            ws_cons.column_dimensions["G"].width = 48

            for nome_aba, df_ind in abas_individuais.items():
                df_ind.to_excel(writer, sheet_name=nome_aba, index=False)
                ws_aba = writer.sheets[nome_aba]
                ws_aba.column_dimensions["A"].width = 32
                ws_aba.column_dimensions["C"].width = 12
                ws_aba.column_dimensions["D"].width = 68
                ws_aba.column_dimensions["E"].width = 16
                ws_aba.column_dimensions["G"].width = 48

        print(f"\n✅ SUCESSO! Relatório gerado em: {caminho_final}")


if __name__ == "__main__":
    processar_revpack()