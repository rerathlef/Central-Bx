import pandas as pd
import os
import re
from datetime import datetime
import warnings
import sys
warnings.simplefilter("ignore")

# =============================================================================
# 1. CONFIGURAÇÕES
# =============================================================================


if len(sys.argv) > 1 and not sys.argv[1].startswith("-f"):
    PASTA_TRABALHO = sys.argv[1]
else:
    # Se você rodar o arquivo sozinho, ele usa o que estiver aqui:
    PASTA_TRABALHO = r"G:\.shortcut-targets-by-id\1NZDdbjsov2A4Nb33L8NlwmII9RQeNMuu\12 Cativa\03 Financeiro\01. Extratos\2026\05 - Maio\22.05.2026\SICOOB"

MAPEAMENTO_PARES = {
    "SICOOB - CATIVA TEXTIL": "COMPROVANTES - CATIVA TEXTIL",
    "SICOOB - CATIVA APIUNA": "COMPROVANTES - CATIVA APIUNA",
    "SICOOB - CARISMA": "COMPROVANTES - CARISMA",
    "SICOOB - CATIVA MS": "COMPROVANTES - CATIVA MS",
    "SICOOB - CDT": "COMPROVANTES - CDT",
    "SICOOB - CATIVA MS HORIZONTE": "COMPROVANTES - CATIVA MS HORIZONTE",
}

MAPEAMENTO_FIXO = {
    "TAR": "BANCO SICOOB",
    "JUROS": "BANCO SICOOB",
    "IOF": "GOVERNO FEDERAL",
    "DÉBITO PACOTE SERVIÇOS": "BANCO SICOOB"
}

# Termos que devem ser ignorados se aparecerem como linha principal (ruído do Sicoob)
TERMOS_IGNORAR = ["RECEBIMENTO PIX", "PAGAMENTO PIX"]

# =============================================================================
# 2. FUNÇÕES DE TRATAMENTO
# =============================================================================

def limpar_valor_sicoob_real(valor):
    if pd.isna(valor) or str(valor).strip() == "": return 0.0
    texto = str(valor).upper().strip().replace("\xa0", "")
    
    # AGORA ELE PROCURA O "D" OU O SINAL DE MENOS "-" PARA SABER SE É NEGATIVO
    multiplicador = -1 if ("D" in texto or "-" in texto) else 1
    
    # A VASSOURA CONTINUA APAGANDO O SINAL DE MENOS, MAS NÃO TEM PROBLEMA 
    # PORQUE JÁ SALVAMOS ESSA INFORMAÇÃO NO MULTIPLICADOR ACIMA
    texto_limpo = re.sub(r'[^0-9,.]', '', texto)
    
    if "," in texto_limpo and "." in texto_limpo:
        texto_limpo = texto_limpo.replace(".", "").replace(",", ".")
    elif "," in texto_limpo:
        texto_limpo = texto_limpo.replace(",", ".")
    try:
        return round(float(texto_limpo) * multiplicador, 2)
    except:
        return 0.0

def tratar_favorecido_g_f(g_val, f_val):
    g_limpo = ""
    if pd.notna(g_val) and str(g_val).strip() != "":
        texto = str(g_val).upper()
        if "/" in texto: texto = texto.split("/")[-1]
        texto = texto.replace("CPF", "").replace("CNPJ", "")
        texto = re.sub(r'\d{3,}', '', texto)
        texto = texto.replace(".", " ").replace("-", " ")
        g_limpo = " ".join(texto.split()).strip()
    if g_limpo == "" and pd.notna(f_val):
        return str(f_val).strip().upper()
    return g_limpo

def data_para_serial_excel(data):
    if pd.isna(data): return None
    try:
        if isinstance(data, str):
            data = pd.to_datetime(data, dayfirst=True)
        base_date = datetime(1899, 12, 30)
        delta = data - base_date
        return delta.days
    except:
        return None

# =============================================================================
# 3. PROCESSAMENTO
# =============================================================================

def conciliador_cativa_clean():
    print(f"--- 🚀 INICIANDO CONCILIAÇÃO: VERSÃO CLEAN (SICOOB) ---")
    arquivos_na_pasta = os.listdir(PASTA_TRABALHO)
    relatorio_consolidado = []

    for termo_ext, termo_comp in MAPEAMENTO_PARES.items():
        arq_ext_list = [f for f in arquivos_na_pasta if f == f"{termo_ext}.xlsx"]
        arq_comp_list = [f for f in arquivos_na_pasta if f == f"{termo_comp}.xlsx"]

        if not arq_ext_list or not arq_comp_list: continue

        nome_unidade = termo_ext.replace("SICOOB - ", "").strip()

        # A. COMPROVANTE (SISTEMA)
        df_comp_raw = pd.read_excel(os.path.join(PASTA_TRABALHO, arq_comp_list[0]), sheet_name=0)
        df_comp = df_comp_raw[df_comp_raw['Situação'].astype(str).str.upper().str.contains("EFETIVA|REALIZA|PAGO", na=False)].copy()
        
        df_comp['DATA_STR'] = pd.to_datetime(df_comp['Data'], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y')
        df_comp['VALOR_ABS'] = df_comp['Valor'].apply(lambda x: abs(float(x)) if pd.notna(x) else 0.0).round(2)

        indices_comprovante_usados = []

        # B. EXTRATO (BANCO)
        df_ext = pd.read_excel(os.path.join(PASTA_TRABALHO, arq_ext_list[0]), skiprows=1, sheet_name=0)
        df_ext.columns = df_ext.columns.str.strip().str.upper()
        
        for i in range(len(df_ext)):
            linha = df_ext.iloc[i]
            
            # PULA LINHAS SEM VALOR OU QUE SÃO APENAS RUÍDO (CNPJ/PAGAMENTO PIX)
            hist_ext = str(linha['HISTÓRICO']).strip()
            if pd.isna(linha.get('VALOR')) or any(term in hist_ext.upper() for term in TERMOS_IGNORAR) or re.search(r'\d{2}\.\d{3}\.\d{3}', hist_ext):
                continue

            valor_ext = limpar_valor_sicoob_real(linha['VALOR'])
            data_ext_str = pd.to_datetime(linha['DATA'], dayfirst=True, errors='coerce').strftime('%d/%m/%Y')

            # --- REGRA DE CRÉDITO (JUNÇÃO) ---
            if valor_ext > 0 and "SALDO" not in hist_ext.upper():
                if i + 1 < len(df_ext) and pd.isna(df_ext.iloc[i+1]['DATA']):
                    hist_ext = f"{hist_ext} - {str(df_ext.iloc[i+1]['HISTÓRICO']).strip()}"
                favorecido, valor_interno, check = nome_unidade, 0.0, "SIM"
            else:
                favorecido, valor_interno, check = "", 0.0, ""

            # --- REGRAS FIXAS ---
            for termo, nome_f in MAPEAMENTO_FIXO.items():
                if termo in hist_ext.upper():
                    favorecido, check = nome_f, "SIM"; break
            
            # REGRA DO SALDO (FAVORECIDO VAZIO)
            if "SALDO" in hist_ext.upper(): 
                favorecido = "" # Removido "SALDO BANCÁRIO"
                check = ""

            # --- MATCH 1 PARA 1 (PADRÃO BR) ---
            if valor_ext < 0 and favorecido == "":
                v_procurado = round(abs(valor_ext), 2)
                candidatos = df_comp[(df_comp['VALOR_ABS'] == v_procurado) & (df_comp['DATA_STR'] == data_ext_str)]
                
                match_sucesso = False
                for idx_comp, row_comp in candidatos.iterrows():
                    if idx_comp not in indices_comprovante_usados:
                        favorecido = tratar_favorecido_g_f(row_comp['Conta Destino'], row_comp['Beneficiário'])
                        valor_interno = row_comp['Valor']
                        check = "SIM"
                        indices_comprovante_usados.append(idx_comp)
                        match_sucesso = True
                        break
                
                if not match_sucesso:
                    check = "NÃO ENCONTRADO"

            relatorio_consolidado.append({
                'AG./ORIGEM': termo_ext,
                " ": " ",
                'DATA': data_para_serial_excel(linha['DATA']),
                'LANÇAMENTO EXTRATO': hist_ext,
                'VALOR (R$)': f"{valor_ext:.2f}".replace(".", ","),
                " ": " ",
                'VALOR INTERNO': f"{abs(float(valor_interno)):.2f}".replace(".", ","),
                'FAVORECIDO': favorecido,
                'CHECK': check
            })

    if relatorio_consolidado:
        df_final = pd.DataFrame(relatorio_consolidado)
        nome = f"AUT C.B - CTV - SICOOB.xlsx"
        df_final.to_excel(os.path.join(PASTA_TRABALHO, nome), index=False)
        print(f"✅ Missão Cumprida: {nome}")

if __name__ == "__main__":
    conciliador_cativa_clean()
