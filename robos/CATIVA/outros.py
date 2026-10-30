import pandas as pd
import os
import re
from datetime import datetime
import warnings
import sys

warnings.simplefilter("ignore")

print("--- 🚀 MÁQUINA DE CONCILIAÇÃO V8 - FILTRO DE LIMPEZA FINAL ---")

# =============================================================================
# 1. CONFIGURAÇÕES E REGRAS
# =============================================================================
# Se o Mestre mandar a pasta, ele usa. Se você rodar ele sozinho, usa a fixa.
if len(sys.argv) > 1:
    PASTA_ARQUIVOS = sys.argv[1]
else:
    PASTA_ARQUIVOS = r"G:\.shortcut-targets-by-id\1NZDdbjsov2A4Nb33L8NlwmII9RQeNMuu\12 Cativa\03 Financeiro\01. Extratos\2026\05 - Maio\06.05.2026"
    
NOME_RELATORIO_FINAL = "CONCILIACAO_CONSOLIDADA_CATIVA.xlsx"

# LISTA DE TERMOS PARA REMOVER (Lixo que vem no extrato do banco)
TERMOS_EXCLUIR = ["S A L D O", "SALDO ANTERIOR", "SALDO INICIAL", "SALDO FINAL", "S.A.L.D.O", "SALDO DO DIA"]

CONFIG_BANCOS = {
    "BMP": {
        "id_arquivo": "BMP",
        "col_data": "Data_da_Ocorrencia",
        "col_desc": "Lancamento",
        "col_favorecido_raw": "Nome",
        "col_saldo": "Saldo",
        "tipo_valor": "UNICA",
        "col_valor_1": "Valor",
        "regra_saldo": "EXTRAIR_TOPO",
        "nome_banco_padrao": "BMP"
    },
    "VORTX": {
        "id_arquivo": "VORTX",
        "col_data": "Data_da_Ocorrencia",
        "col_desc": "Lancamento",
        "col_favorecido_raw": "Nome",
        "col_saldo": "Saldo",
        "tipo_valor": "UNICA",
        "col_valor_1": "Valor",
        "regra_saldo": "EXTRAIR_TOPO",
        "nome_banco_padrao": "VORTX"
    },
    "BB": {
        "id_arquivo": "BB",
        "col_data": "Data",
        "col_desc": "Historico",         
        "col_favorecido_raw": "Detalhamento Hist.",
        "tipo_valor": "DC", 
        "col_valor_1": "Valor R$",
        "col_sinal": "Inf.", 
        "regra_saldo": "CALCULAR_ACUMULADO",
        "nome_banco_padrao": "BB"
    }
}

REGRAS_AVANCADAS = [
    {"termo": "TARIFA", "tipo": "FIXO_BANCO"},
    {"termo": "TAR ", "tipo": "FIXO_BANCO"},
    {"termo": "RECEBIMENTO BOLETO", "tipo": "DONO_CONTA"},
    {"termo": "Recebimento de boletos", "tipo": "DONO_CONTA"},
    {"termo": "Recebimento de boletos", "tipo": "DONO_CONTA"},
    {"termo": "LIQUIDACAO BOLETO", "tipo": "DONO_CONTA"},
    {"termo": "PIX RECEBIDO", "tipo": "DONO_CONTA"},
    {"termo": "IOF", "tipo": "FIXO", "valor": "RECEITA FEDERAL"},
]

# =============================================================================
# 2. FUNÇÕES DE APOIO
# =============================================================================

def limpar_valor_moeda(valor):
    if pd.isna(valor) or str(valor).strip() == "": return 0.0
    v = str(valor).strip().replace("R$", "").replace(" ", "").replace('"', '').replace("'", "")
    if "," in v and "." in v: v = v.replace(".", "").replace(",", ".")
    elif "," in v: v = v.replace(",", ".")
    try: return round(float(v), 2)
    except: return 0.0

def date_to_excel_serial(date_val):
    if pd.isna(date_val): return ""
    try: return (date_val - datetime(1899, 12, 30)).days
    except: return ""

def identificar_origem_completa(nome_arquivo):
    nome = os.path.splitext(nome_arquivo)[0].upper()
    if " - " in nome:
        partes = nome.split(" - ", 1)
        return partes[0].strip(), partes[1].strip()
    return nome, "EMPRESA"

def encontrar_cabecalho(caminho, nome_coluna_data):
    try:
        df_temp = pd.read_excel(caminho, nrows=40, header=None)
        for i, row in df_temp.iterrows():
            linha_texto = [str(x).strip().upper() for x in row.values]
            if str(nome_coluna_data).upper() in linha_texto: return i
    except: pass
    return None

# =============================================================================
# 3. MOTOR DE PROCESSAMENTO
# =============================================================================

def processar():
    if not os.path.exists(PASTA_ARQUIVOS):
        print(f"❌ Pasta não encontrada!"); return

    arquivos = [f for f in os.listdir(PASTA_ARQUIVOS) if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith("~$")]
    
    # --- ADICIONE ESTA TRAVA DE AVISO AQUI ---
    if len(arquivos) == 0:
        print("⚠️ [AVISO] Nenhum extrato bancário encontrado na pasta de OUTROS BANCOS para esta data.")
        return
    # ----------------------------------------

    dados_consolidados = []
    abas_individuais = {}
    
    # ... (o resto do código continua igualzinho para baixo)

    for arquivo in arquivos:
        print(f"➜ Processando: {arquivo}")
        banco_arq, empresa_arq = identificar_origem_completa(arquivo)
        config = next((cfg for key, cfg in CONFIG_BANCOS.items() if key in banco_arq), None)
        if not config: continue

        try:
            caminho = os.path.join(PASTA_ARQUIVOS, arquivo)
            cabecalho = encontrar_cabecalho(caminho, config["col_data"])
            df_original = pd.read_excel(caminho, skiprows=cabecalho)
            df_original.columns = [str(c).strip() for c in df_original.columns]

            # 1. Captura de Saldo Anterior (Exclusivo para o BB - ANTES DA LIMPEZA)
            saldo_acumulado = 0.0
            if config.get("regra_saldo") == "CALCULAR_ACUMULADO":
                linha_saldo_ant = df_original[df_original[config["col_desc"]].astype(str).str.upper().str.contains("SALDO ANTERIOR")]
                if not linha_saldo_ant.empty:
                    saldo_acumulado = limpar_valor_moeda(linha_saldo_ant.iloc[0].get(config["col_valor_1"], 0))

            # 2. Tratamento de Transações e FILTRO DE LIXO
            df_original["DATA_OBJ"] = pd.to_datetime(df_original[config["col_data"]], dayfirst=True, errors='coerce')
            df_processado = df_original.dropna(subset=["DATA_OBJ"]).copy()
            
            # Remove linhas que contêm termos de saldo originais do banco
            padrao_exclusao = "|".join(TERMOS_EXCLUIR)
            df_processado = df_processado[~df_processado[config["col_desc"]].astype(str).str.upper().str.contains(padrao_exclusao, na=False)]

            # 3. Valores e Favorecidos
            if config["tipo_valor"] == "DC":
                col_v = [c for c in df_processado.columns if config["col_valor_1"] in c][0]
                df_processado["VALOR_FINAL"] = df_processado.apply(lambda r: round(-abs(limpar_valor_moeda(r[col_v])), 2) if "D" in str(r.get(config["col_sinal"], "")).upper() else round(abs(limpar_valor_moeda(r[col_v])), 2), axis=1)
            else:
                df_processado["VALOR_FINAL"] = df_processado[config["col_valor_1"]].apply(limpar_valor_moeda)

            # (Regras de Favorecido)
            favs = []
            for _, row in df_processado.iterrows():
                desc = str(row[config["col_desc"]]).upper()
                fav_raw = str(row[config.get("col_favorecido_raw", config["col_desc"])]).upper()
                final_fav = fav_raw
                regra_on = False
                for regra in REGRAS_AVANCADAS:
                    if regra["termo"] in desc:
                        if regra["tipo"] == "FIXO_BANCO": final_fav = config["nome_banco_padrao"]
                        elif regra["tipo"] == "DONO_CONTA": final_fav = empresa_arq
                        elif regra["tipo"] == "FIXO": final_fav = regra["valor"]
                        regra_on = True; break
                if not regra_on and config["id_arquivo"] == "BB":
                    final_fav = re.sub(r'^[\d/:\s\.-]+', '', fav_raw).strip()
                favorecidos_finais = favs.append(final_fav)
            df_processado["FAV_FINAL"] = favs

            # 4. DF Base
            df_aba = pd.DataFrame({
                "AG./ORIGEM": f"{banco_arq} - {empresa_arq}", " ": "",
                "DATA_SERIAL": df_processado["DATA_OBJ"].apply(date_to_excel_serial),
                "DATA": df_processado["DATA_OBJ"].apply(date_to_excel_serial),
                "LANÇAMENTO EXTRATO": df_processado[config["col_desc"]],
                "VALOR (R$)": df_processado["VALOR_FINAL"],
                "  ": "", "FAVORECIDO": df_processado["FAV_FINAL"], "CHECK": "",
                "priority": 0, "original_index": df_processado.index
            })

            # 5. Regras de Saldo
            saldos_fechamento = []
            if config.get("regra_saldo") == "EXTRAIR_TOPO":
                for data_serial in df_aba["DATA_SERIAL"].unique():
                    subset_dia = df_aba[df_aba["DATA_SERIAL"] == data_serial]
                    idx_fechamento = subset_dia["original_index"].min()
                    valor_saldo = round(limpar_valor_moeda(df_original.loc[idx_fechamento, config["col_saldo"]]), 2)
                    saldos_fechamento.append({"AG./ORIGEM": f"{banco_arq} - {empresa_arq}", "DATA": data_serial, "LANÇAMENTO EXTRATO": "SALDO FINAL", "VALOR (R$)": valor_saldo, "priority": 1, "DATA_SERIAL": data_serial})

            elif config.get("regra_saldo") == "CALCULAR_ACUMULADO":
                df_aba = df_aba.sort_values(by=["DATA_SERIAL", "original_index"])
                for data_serial in df_aba["DATA_SERIAL"].unique():
                    movimentacao_dia = round(df_aba[df_aba["DATA_SERIAL"] == data_serial]["VALOR (R$)"].sum(), 2)
                    saldo_acumulado = round(saldo_acumulado + movimentacao_dia, 2)
                    saldos_fechamento.append({"AG./ORIGEM": f"{banco_arq} - {empresa_arq}", "DATA": data_serial, "LANÇAMENTO EXTRATO": "SALDO FINAL", "VALOR (R$)": saldo_acumulado, "priority": 1, "DATA_SERIAL": data_serial})

            df_aba = pd.concat([df_aba, pd.DataFrame(saldos_fechamento)], ignore_index=True)
            df_aba = df_aba.sort_values(by=["DATA_SERIAL", "priority"]).drop(columns=["priority", "original_index", "DATA_SERIAL"]).fillna("")

            if not df_aba.empty:
                dados_consolidados.append(df_aba)
                abas_individuais[f"{banco_arq}-{empresa_arq}"[:30]] = df_aba

        except Exception as e:
            print(f"   ❌ Erro em {arquivo}: {e}")

    # =============================================================================
    # 4. O LIMPA NO RELATÓRIO FINAL (A MÁGICA FINAL)
    # =============================================================================
    if dados_consolidados:
        df_completo = pd.concat(dados_consolidados, ignore_index=True)

        # 1. Remove linhas onde o Lançamento é vazio ou apenas espaços
        df_completo = df_completo.dropna(subset=["LANÇAMENTO EXTRATO"])
        df_completo = df_completo[df_completo["LANÇAMENTO EXTRATO"].astype(str).str.strip() != ""]

        # 2. Filtro de Segurança: Garante que nenhum termo da lista TERMOS_EXCLUIR passou
        # MAS: Mantemos o nosso "SALDO FINAL" (nós criamos ele, ele é o oficial)
        padrao_limpa = "|".join([t for t in TERMOS_EXCLUIR if t != "SALDO FINAL"])
        # Filtra apenas linhas que NÃO são o nosso saldo final, mas que contêm termos proibidos
        mask_lixo = (df_completo["LANÇAMENTO EXTRATO"].str.upper().str.contains(padrao_limpa, na=False)) & \
                    (df_completo["LANÇAMENTO EXTRATO"].str.upper() != "SALDO FINAL")
        
        df_completo = df_completo[~mask_lixo]

        # Salva o resultado
        caminho_final = os.path.join(PASTA_ARQUIVOS, NOME_RELATORIO_FINAL)
        with pd.ExcelWriter(caminho_final, engine='openpyxl') as writer:
            df_completo.to_excel(writer, sheet_name='CONSOLIDADO', index=False)
            for nome, df_ind in abas_individuais.items(): 
                # Aplica o mesmo limpa nas abas individuais
                mask_lixo_aba = (df_ind["LANÇAMENTO EXTRATO"].str.upper().str.contains(padrao_limpa, na=False)) & \
                                (df_ind["LANÇAMENTO EXTRATO"].str.upper() != "SALDO FINAL")
                df_ind[~mask_lixo_aba].to_excel(writer, sheet_name=nome, index=False)
        
        print(f"\n✅ SUCESSO! Relatório limpo e consolidado gerado.")

if __name__ == "__main__":
    processar()
