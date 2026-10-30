import pandas as pd
import os
import datetime
import warnings
import time
import re
import ctypes  # Para janelas de alerta do Windows

# Ignora avisos de formatação do Excel
warnings.filterwarnings('ignore')

# ==============================================================================
# PAINEL DE CONTROLE - CONFIGURAÇÃO DE PASTAS E VIGIA
# ==============================================================================

# --- COLOQUE O ENDEREÇO DA PASTA DOS ARQUIVOS AQUI ---
import sys
if len(sys.argv) > 1 and not sys.argv[1].startswith("-"):
    PASTA_FONTE = sys.argv[1]
else:
    PASTA_FONTE = r'G:\.shortcut-targets-by-id\1HH8YPWOEhDI7x9lk-LNQXlVf9sUusDFg\08 Revpack\03 Financeiro\02. EXTRATOS\POSIÇÕES FUNDOS\2026\10 - OUT\Fundos 13.10\Trabalhados\teste robo 1'
PASTA_DESTINO = PASTA_FONTE

# --- EMPRESA FIXA (GRUPO ÚNICO) ---
EMPRESA_FIXA = 'REVPACK'

# --- NOMES DOS ARQUIVOS GERADOS ---
NOME_ARQUIVO_FINAL = 'CONTROLE DE SACADOS AUTOMATIZADO - REVPACK.xlsx'
NOME_ARQUIVO_DUPLICADOS = 'CONTROLE DUPLICADOS - REVPACK.xlsx'
NOME_ARQUIVO_LOG_ERROS = 'LOG_ERROS - REVPACK.xlsx'

# --- ATIVAR O MODO VIGIA? ---
# True = O robô fica aguardando todos os arquivos da lista chegarem na pasta.
# False = O robô roda imediatamente uma única vez com os arquivos que estiverem lá.
ATIVAR_MODO_VIGIA = False

# --- QUAIS ARQUIVOS DEVEM ESTAR NA PASTA PARA O ROBÔ RODAR? ---
ARQUIVOS_ESPERADOS_VIGIA = [
    "911",
    "ATHENA",
    "BANKOK",
    "C6",
    "CAPITAL FINANÇAS",
    "CONTINENTAL",
    "DAYCOVAL",
    "INVISTA"
]

# ==============================================================================
# MAPEAMENTOS DOS FUNDOS (REVPACK)
# ==============================================================================
MAPEAMENTOS = {
    '911': {
        'header_row': 1,
        'filtro': {'coluna': 'Sacado', 'acao': 'remover_vazios'},
        'SACADO': 'Sacado',
        'DOCUMENTO': 'Documento',
        'DATA DE VENCIMENTO': 'Vencimento',
        'VALOR FACE': 'Valor',
        'MODO DE COBRANÇA': 'IDG',
        'BORDERO': 'Ope',
        'NOME_FUNDO_FIXO': '911'
    },
    'ATHENA': {
        'header_row': 0,
        'filtro': {'coluna': 'Vencto', 'acao': 'remover_vazios'},
        'SACADO': 'Nome/Razão Sacado',
        'DOCUMENTO': 'Titulo',
        'DATA DE VENCIMENTO': 'Vencto',
        'VALOR FACE': 'Valor',
        'MODO DE COBRANÇA': 'Tipo',
        'BORDERO': 'Bordero',
        'NOME_FUNDO_FIXO': 'ATHENA'
    },
    'CAPITAL FINANÇAS': {
        'header_row': 0,
        'filtro': {'coluna': 'Vencto', 'acao': 'remover_vazios'},
        'SACADO': 'Nome/Razão Sacado',
        'DOCUMENTO': 'Titulo',
        'DATA DE VENCIMENTO': 'Vencto',
        'VALOR FACE': 'Valor',
        'MODO DE COBRANÇA': 'Tipo',
        'BORDERO': 'Bordero',
        'NOME_FUNDO_FIXO': 'CAPITAL FINANÇAS'
    },
    'CAPITAL FINANCAS': {  # Caso salvem sem cedilha
        'header_row': 0,
        'filtro': {'coluna': 'Vencto', 'acao': 'remover_vazios'},
        'SACADO': 'Nome/Razão Sacado',
        'DOCUMENTO': 'Titulo',
        'DATA DE VENCIMENTO': 'Vencto',
        'VALOR FACE': 'Valor',
        'MODO DE COBRANÇA': 'Tipo',
        'BORDERO': 'Bordero',
        'NOME_FUNDO_FIXO': 'CAPITAL FINANÇAS'
    },
    'CONTINENTAL': {
        'header_row': 0,
        'filtro': {'coluna': 'Vencto', 'acao': 'remover_vazios'},
        'SACADO': 'Nome/Razão Sacado',
        'DOCUMENTO': 'Titulo',
        'DATA DE VENCIMENTO': 'Vencto',
        'VALOR FACE': 'Valor',
        'MODO DE COBRANÇA': 'Tipo',
        'BORDERO': 'Bordero',
        'NOME_FUNDO_FIXO': 'CONTINENTAL'
    },
    'C6': {
        'header_row': 0,
        'filtro': {'coluna': 'Venc', 'acao': 'remover_vazios'},
        'SACADO': 'Sacado',
        'DOCUMENTO': 'Documento',
        'DATA DE VENCIMENTO': 'Venc',
        'VALOR FACE': 'Valor',
        'MODO DE COBRANÇA': 'IDG',
        'BORDERO': 'Operação',
        'NOME_FUNDO_FIXO': 'C6'
    },
    'INVISTA': {
        'header_row': 0,
        'filtro': {'coluna': 'Venc', 'acao': 'remover_vazios'},
        'SACADO': 'Sacado',
        'DOCUMENTO': 'Documento',
        'DATA DE VENCIMENTO': 'Venc',
        'VALOR FACE': 'Valor',
        'MODO DE COBRANÇA': 'IG',
        'BORDERO': 'Operação',
        'NOME_FUNDO_FIXO': 'INVISTA'
    },
    'DAYCOVAL': {
        'header_row': 6,
        'formato_data': 'extrair_final',
        'filtro': {'coluna': 'Sacado', 'acao': 'remover_vazios'},
        'SACADO': 'Sacado',
        'DOCUMENTO': 'Seu Número',
        'DATA DE VENCIMENTO': 'Vencimento',
        'VALOR FACE': 'Valor',
        'MODO DE COBRANÇA': 'Nome do Produto',
        'NOME_FUNDO_FIXO': 'DAYCOVAL'
    },
    'BANKOK': {
        'header_row': 0,
        # Filtra apenas títulos que têm saldo em "A Receber" > 0 (ignora os liquidados)
        'filtro': {'coluna': 'A Receber', 'acao': 'maior_que_zero'},
        'SACADO': 'Pagador',
        'DOCUMENTO': 'NDoc',
        'DATA DE VENCIMENTO': 'Vencimento',
        'VALOR FACE': 'Vl_Face',
        'MODO DE COBRANÇA': 'TDoc',
        'BORDERO': 'Operação',
        'NOME_FUNDO_FIXO': 'BANKOK'
    },
}

# ==============================================================================
# SUPER DICIONÁRIO DE CORREÇÕES - MODO DE COBRANÇA
# ==============================================================================
correcoes_modo_cobranca = {
    # --------------------------------------------------------------------------
    # ⚠️ PENDENTES REVPACK (SUBSTITUIR PELO MODO CORRETO APÓS VALIDAR COM A EQUIPE)
    # --------------------------------------------------------------------------
    '326011': '326011 (DEFINIR C6)',
    '326017': '326017 (DEFINIR C6)',
    '0326011': '326011 (DEFINIR C6)',
    '0326017': '326017 (DEFINIR C6)',
    '10701': '10701 (DEFINIR INVISTA)',
    '010701': '10701 (DEFINIR INVISTA)',
    '10901': '10901 (DEFINIR 911)',
    '010901': '10901 (DEFINIR 911)',
    '10103': '10103 (DEFINIR 911)',
    '010103': '10103 (DEFINIR 911)',
    'DE': 'DESCONTO',  # Caso use a coluna 'Mod' no Bankok

    # --------------------------------------------------------------------------
    # DESCONTO (CONSOLIDADO ESSERE + CATIVA + WETZEL)
    # --------------------------------------------------------------------------
    'FD-DP': 'DESCONTO', 'DES': 'DESCONTO', 'CS-DP': 'DESCONTO', 'DUP (D-01)': 'DESCONTO',
    'DUP': 'DESCONTO', 'FIDC2': 'DESCONTO', 'FIDC': 'DESCONTO', 'DP': 'DESCONTO',
    'DMR': 'DESCONTO', 'CESSÃO RECEBIVEIS COM COOBRIGAÇÃO': 'DESCONTO', 'C. PRÓPRIA': 'DESCONTO',
    'C. Própria': 'DESCONTO', 'D': 'DESCONTO', 'DM': 'DESCONTO', 'FIDCNPRJ': 'DESCONTO',
    'DESCONTO': 'DESCONTO', 'Desconto': 'DESCONTO', 'DUPLICATA': 'DESCONTO', '220152': 'DESCONTO',
    'DESC-EMPRESTIMOS DESCONTO': 'DESCONTO', '010101 / Desconto - Duplicata - Operado': 'DESCONTO',
    '010101 / DESCONTO - DUPLICATA - OPERADO': 'DESCONTO',
    '010192 / Desconto - Duplicata - Sacado DDA': 'DESCONTO',
    '010192 / DESCONTO - DUPLICATA - SACADO DDA': 'DESCONTO',
    '10192': 'DESCONTO', '010192': 'DESCONTO', '14601': 'DESCONTO', 'CD': 'DESCONTO',
    'CH': 'DESCONTO', '010101': 'DESCONTO', '010601': 'DESCONTO', '010630': 'DESCONTO',
    '010631': 'DESCONTO',

    # --------------------------------------------------------------------------
    # COMISSÁRIA
    # --------------------------------------------------------------------------
    'FD-CM': 'COMISSÁRIA', 'CM': 'COMISSÁRIA', 'Comissária': 'COMISSÁRIA',
    'COMISSÁRIA': 'COMISSÁRIA', 'DUPLICATA MERCANTIL': 'COMISSÁRIA', '030101': 'COMISSÁRIA',

    # --------------------------------------------------------------------------
    # GARANTIA
    # --------------------------------------------------------------------------
    'CS-OG': 'GARANTIA', 'SRM - GARKGCF': 'GARANTIA', 'COBRANCA SIMPLES': 'GARANTIA',
    'COBRANCA VINCULADA': 'GARANTIA', '10101': 'GARANTIA', 'Cobrança Vinculada': 'GARANTIA',
    'COBRANÇA VINCULADA': 'GARANTIA', 'Cobrança Simples': 'GARANTIA', 'COBRANÇA SIMPLES': 'GARANTIA',
    'RCR -RAPIDA COM REGISTRO': 'GARANTIA', '220101 / Trustee - Duplicata - Operado': 'GARANTIA',
    '220101 / TRUSTEE - DUPLICATA - OPERADO': 'GARANTIA',

    # --------------------------------------------------------------------------
    # FOMENTO
    # --------------------------------------------------------------------------
    'FD-CA': 'FOMENTO', 'FOM': 'FOMENTO', 'DS': 'FOMENTO', 'FD-IN': 'FOMENTO',
    'NP': 'FOMENTO', 'FLUXO FUTURO FIDC': 'FOMENTO', 'PD': 'FOMENTO', 'CHQ': 'FOMENTO',

    # --------------------------------------------------------------------------
    # NOTA COMERCIAL
    # --------------------------------------------------------------------------
    'FD-CT': 'NOTA COMERCIAL', 'FD-NC': 'NOTA COMERCIAL', '220157': 'NOTA COMERCIAL',
    '220101': 'NOTA COMERCIAL', 'NOTAS PROMISSÓRIAS': 'NOTA COMERCIAL', 'GARKGCP': 'NOTA COMERCIAL',
    '26201': 'NOTA COMERCIAL', 'NC': 'NOTA COMERCIAL', '090101': 'NOTA COMERCIAL',

    # --------------------------------------------------------------------------
    # CONFISSÃO DE DÍVIDA
    # --------------------------------------------------------------------------
    'RD-DP': 'CONFISSÃO DE DIVIDA',
}

# ==============================================================================
# FUNÇÕES AUXILIARES E LEITOR CAMALEÃO (.XLSX, .XLS, .CSV)
# ==============================================================================

def exibir_popup(titulo, mensagem, estilo=64):
    """Exibe um popup do Windows (Estilo 64 = Info, 16 = Erro)"""
    try:
        ctypes.windll.user32.MessageBoxW(0, mensagem, titulo, estilo)
    except Exception:
        print(f"[{titulo}] {mensagem}")

def verificar_arquivos_presentes():
    """Verifica se todos os arquivos obrigatórios estão na pasta."""
    try:
        arquivos_na_pasta = os.listdir(PASTA_FONTE)
        arquivos_na_pasta_upper = [f.upper() for f in arquivos_na_pasta if not f.startswith("~")]
    except Exception as e:
        print(f"Erro ao ler pasta: {e}")
        return False, []

    faltantes = []
    for esperado in ARQUIVOS_ESPERADOS_VIGIA:
        esperado_upper = esperado.upper()
        encontrou = any(esperado_upper in arq_real for arq_real in arquivos_na_pasta_upper)
        if not encontrou:
            faltantes.append(esperado)

    return (len(faltantes) == 0), sorted(faltantes)

def ler_arquivo_universal(caminho_arquivo, header_linha=0):
    """
    Leitor blindado para .xlsx, .xls (antigo 97-2003 ou HTML disfarçado) e .csv.
    Se faltar o pacote 'xlrd' no Anaconda, converte em segundo plano via Excel (win32com).
    """
    ext = os.path.splitext(caminho_arquivo)[1].lower()

    if ext == '.xlsx':
        return pd.read_excel(caminho_arquivo, sheet_name=0, header=header_linha, engine='openpyxl')

    elif ext == '.csv':
        try:
            return pd.read_csv(caminho_arquivo, sep=';', decimal=',', encoding='latin1', header=header_linha)
        except Exception:
            return pd.read_csv(caminho_arquivo, sep=',', decimal='.', encoding='latin1', header=header_linha)

    elif ext == '.xls':
        # 1ª Tentativa: Motor nativo xlrd
        try:
            return pd.read_excel(caminho_arquivo, sheet_name=0, header=header_linha, engine='xlrd')
        except Exception as e_xlrd:
            # 2ª Tentativa: Usa o próprio Excel instalado no Windows para converter temporariamente
            try:
                import win32com.client
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                caminho_abs = os.path.abspath(caminho_arquivo)
                caminho_temp = caminho_abs + "_temp_conv.xlsx"
                wb = excel.Workbooks.Open(caminho_abs)
                wb.SaveAs(caminho_temp, FileFormat=51)  # 51 = .xlsx
                wb.Close()
                excel.Quit()
                df = pd.read_excel(caminho_temp, sheet_name=0, header=header_linha, engine='openpyxl')
                if os.path.exists(caminho_temp):
                    os.remove(caminho_temp)
                return df
            except Exception:
                # 3ª Tentativa: Caso seja um HTML/TSV salvo como .xls
                try:
                    dfs = pd.read_html(caminho_arquivo, header=header_linha)
                    if dfs:
                        return dfs[0]
                except Exception:
                    pass
                raise ValueError(
                    f"Não foi possível abrir o arquivo .xls ({os.path.basename(caminho_arquivo)}). "
                    f"Dica: rode 'pip install xlrd' no prompt do Anaconda. Erro original: {e_xlrd}"
                )

def limpar_moeda_br(x):
    if pd.isna(x):
        return pd.NA
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).replace('R$', '').replace(' ', '').strip()
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    return pd.to_numeric(s, errors='coerce')

def limpar_nome_sacado(nome_str):
    """Remove números, códigos de cliente e traços, deixando apenas letras e espaços."""
    if pd.isna(nome_str):
        return pd.NA
    nome = str(nome_str).strip().upper()
    nome = re.sub(r'[^A-ZÀ-Ÿ\s]', ' ', nome)
    nome = re.sub(r'\s+', ' ', nome).strip()
    return nome if nome else pd.NA

def formatar_documento_universal_legado(doc_original):
    if pd.isna(doc_original):
        return doc_original
    doc_str = str(doc_original).strip()
    # Remove sufixo de float (.0) caso o Excel tenha lido número como decimal
    doc_str = re.sub(r'\.0$', '', doc_str)

    if ' ' in doc_str or '-' in doc_str:
        doc_str = doc_str.replace(' ', '/').replace('-', '/')

    # Remove letras de parcela no final (ex: Daycoval '078580/A' -> '078580')
    doc_str = re.sub(r'/[A-Za-z]+$', '', doc_str).strip()

    if '/' in doc_str:
        partes = doc_str.split('/')
        try:
            if len(partes) > 1 and partes[-1].isdigit():
                partes[-1] = str(int(partes[-1]))
                doc_str = '/'.join(partes)
        except ValueError:
            pass
        return doc_str

    if doc_str.isdigit() and len(doc_str) >= 7:
        try:
            prefixo = doc_str[:-3]
            sufixo = doc_str[-3:]
            sufixo_int = int(sufixo)
            if sufixo_int == 0:
                return doc_str
            return f"{prefixo}/{sufixo_int}"
        except Exception:
            return doc_str
    return doc_str

# ==============================================================================
# MOTOR PRINCIPAL DE PROCESSAMENTO
# ==============================================================================

def processar_arquivos_na_pasta():
    print("\n" + "=" * 65)
    print("🚀 INICIANDO PROCESSAMENTO - CONTROLE DE SACADOS REVPACK")
    print(f"Origem:  {PASTA_FONTE}")
    print(f"Destino: {PASTA_DESTINO}")
    print("=" * 65 + "\n")

    arquivo_destino_completo = os.path.join(PASTA_DESTINO, NOME_ARQUIVO_FINAL)

    try:
        arquivos_na_pasta = os.listdir(PASTA_FONTE)
    except Exception as e:
        print(f"ERRO CRÍTICO ao acessar a pasta de origem: {e}")
        return False

    extensoes_validas = ('.xlsx', '.xls', '.csv')
    arquivos_ignorados = {NOME_ARQUIVO_FINAL, NOME_ARQUIVO_DUPLICADOS, NOME_ARQUIVO_LOG_ERROS}

    arquivos_fonte = [
        f for f in arquivos_na_pasta
        if f.lower().endswith(extensoes_validas)
        and f not in arquivos_ignorados
        and not f.startswith('~')
        and not f.endswith('_temp_conv.xlsx')
    ]

    print(f"Encontrados {len(arquivos_fonte)} arquivos para processar.")

    arquivos_com_sucesso = []
    log_de_erros = []
    lista_de_dados = []

    # Ordena chaves da maior para a menor para priorizar nomes compostos (ex: 'CAPITAL FINANÇAS')
    chaves_ordenadas = sorted(MAPEAMENTOS.keys(), key=len, reverse=True)

    for arquivo in arquivos_fonte:
        print("-" * 45)
        print(f"Processando: {arquivo}")
        try:
            caminho_arquivo_leitura = os.path.join(PASTA_FONTE, arquivo)
            nome_base_upper = os.path.splitext(os.path.basename(arquivo))[0].strip().upper()

            # Identifica o fundo procurando a chave dentro do nome do arquivo
            fundo_chave = None
            for chave in chaves_ordenadas:
                if chave in nome_base_upper:
                    fundo_chave = chave
                    break

            if not fundo_chave:
                raise ValueError(f"Nenhum fundo mapeado corresponde ao arquivo '{arquivo}'.")

            mapeamento = MAPEAMENTOS[fundo_chave]
            fundo = mapeamento.get('NOME_FUNDO_FIXO', fundo_chave)
            header_linha = mapeamento.get('header_row', 0)

            # Leitura universal (.xlsx, .xls, .csv)
            df_fonte = ler_arquivo_universal(caminho_arquivo_leitura, header_linha=header_linha)

            # Limpa quebras de linha e espaços extras nos nomes das colunas originais
            df_fonte.columns = [str(c).replace('\n', ' ').strip() for c in df_fonte.columns]

            # Aplicação de filtros configurados no mapeamento
            if 'filtro' in mapeamento:
                filtro_info = mapeamento['filtro']
                col_filtro = filtro_info['coluna'].strip()
                acao = filtro_info.get('acao', 'remover_vazios')

                if col_filtro in df_fonte.columns:
                    if acao == 'remover_vazios':
                        df_fonte.dropna(subset=[col_filtro], inplace=True)
                    elif acao == 'maior_que_zero':
                        valores_num = df_fonte[col_filtro].apply(limpar_moeda_br)
                        df_fonte = df_fonte[valores_num > 0]
                    elif acao == 'manter_apenas':
                        pattern = '|'.join(filtro_info.get('valores', []))
                        df_fonte = df_fonte[df_fonte[col_filtro].astype(str).str.contains(pattern, na=False, case=False)]
                    elif acao == 'excluir_se_contiver':
                        df_fonte = df_fonte[~df_fonte[col_filtro].astype(str).str.contains(filtro_info.get('excluir_se_contiver', ''), na=False, case=False)]

            # Mapeamento das colunas para o DataFrame padronizado
            df_novo = pd.DataFrame()
            chaves_config = {
                'filtro', 'header_row', 'formato_data', 'MODO_FIXO',
                'SEPARAR_MODO_E_DOC', 'NOME_FUNDO_FIXO', 'BORDERO'
            }

            for col_final, col_fonte in mapeamento.items():
                if col_final not in chaves_config:
                    col_fonte_limpa = str(col_fonte).strip()
                    if col_fonte_limpa in df_fonte.columns:
                        df_novo[col_final] = df_fonte[col_fonte_limpa]

            if 'MODO_FIXO' in mapeamento:
                df_novo['MODO DE COBRANÇA'] = mapeamento['MODO_FIXO']

            if 'BORDERO' in mapeamento:
                col_bordero = mapeamento['BORDERO'].strip()
                if col_bordero in df_fonte.columns:
                    df_novo['BORDERO'] = df_fonte[col_bordero].astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '')

            # Tratamento de Data (Extração + Conversão para Serial Excel)
            if mapeamento.get('formato_data') == 'extrair_final' and 'DATA DE VENCIMENTO' in df_novo.columns:
                datas_extraidas = df_novo['DATA DE VENCIMENTO'].astype(str).str.extract(r'(\d{1,2}/\d{1,2}/\d{4})')
                df_novo['DATA DE VENCIMENTO'] = datas_extraidas[0]

            df_novo['DATA DE VENCIMENTO'] = pd.to_datetime(df_novo['DATA DE VENCIMENTO'], errors='coerce', dayfirst=True)
            excel_date_offset = pd.Timestamp('1899-12-30')
            df_novo['DATA DE VENCIMENTO'] = (df_novo['DATA DE VENCIMENTO'] - excel_date_offset).dt.days

            df_novo['EMPRESA'] = EMPRESA_FIXA
            df_novo['FUNDO'] = fundo

            lista_de_dados.append(df_novo)
            arquivos_com_sucesso.append(arquivo)
            print(f"  -> SUCESSO ({len(df_novo)} linhas lidas para o fundo {fundo})")

        except Exception as e:
            msg_erro = str(e)
            print(f"  -> ERRO: {msg_erro}")
            log_de_erros.append({'Arquivo': arquivo, 'Erro': msg_erro})

    # ==========================================================================
    # CONSOLIDAÇÃO E REGRAS FINAIS
    # ==========================================================================
    if len(lista_de_dados) > 0:
        df_final = pd.concat(lista_de_dados, ignore_index=True)
        colunas_ordenadas = [
            'EMPRESA', 'FUNDO', 'SACADO', 'DOCUMENTO',
            'DATA DE VENCIMENTO', 'VALOR FACE', 'MODO DE COBRANÇA', 'BORDERO'
        ]
        df_final = df_final.reindex(columns=colunas_ordenadas, fill_value='')

        print("\n" + "-" * 45)
        print("Aplicando higienização e regras de negócio...")

        # 1. Limpeza de Valor Face
        df_final['VALOR FACE'] = df_final['VALOR FACE'].apply(limpar_moeda_br)
        df_final['VALOR FACE'] = pd.to_numeric(df_final['VALOR FACE'], errors='coerce')
        df_final.dropna(subset=['VALOR FACE', 'DATA DE VENCIMENTO'], inplace=True)

        # 2. Remoção de linhas de Totalização antes de limpar o Sacado
        sacado_bruto = df_final['SACADO'].astype(str).str.upper()
        df_final = df_final[~sacado_bruto.str.contains("TOTAL", na=False)]

        # 3. Limpeza de Números/Códigos no Nome do Sacado (Deixa apenas letras)
        df_final['SACADO'] = df_final['SACADO'].apply(limpar_nome_sacado)
        df_final.dropna(subset=['SACADO'], inplace=True)

        # 4. Padronização do Modo de Cobrança
        df_final['MODO DE COBRANÇA'] = (
            df_final['MODO DE COBRANÇA']
            .astype(str)
            .str.strip()
            .str.replace(r'\.0$', '', regex=True)
            .replace(correcoes_modo_cobranca)
        )

        # 5. Formatação de Documentos (Regra Essere + corte do '/A' + remoção de zeros à esquerda)
        print("  -> Formatando documentos...")
        df_final['DOCUMENTO'] = df_final['DOCUMENTO'].apply(formatar_documento_universal_legado)
        df_final['DOCUMENTO'] = (
            df_final['DOCUMENTO']
            .astype(str)
            .apply(lambda x: '/'.join([p.lstrip('0') or '0' for p in x.split('/')]) if '/' in x else (x.lstrip('0') or '0'))
        )

        # 6. Relatório de Duplicatas (Documento + Valor Face)
        print("  -> Verificando duplicatas (Documento + Valor Face)...")
        duplicatas_mask = df_final.duplicated(subset=['DOCUMENTO', 'VALOR FACE'], keep=False)
        df_duplicados = df_final[duplicatas_mask].sort_values(by=['DOCUMENTO'])
        arquivo_duplicados_completo = os.path.join(PASTA_DESTINO, NOME_ARQUIVO_DUPLICADOS)

        try:
            if not df_duplicados.empty:
                df_duplicados.to_excel(arquivo_duplicados_completo, sheet_name='Duplicados', index=False)
                print(f"     [OK] '{NOME_ARQUIVO_DUPLICADOS}' gerado ({len(df_duplicados)} linhas).")
            else:
                print("     [INFO] Nenhuma duplicata encontrada.")
        except Exception as e:
            print(f"     [ERRO] Ao salvar duplicatas: {e}")

        # 7. Salvamento do Arquivo Principal
        try:
            df_final.to_excel(arquivo_destino_completo, sheet_name='BASE', index=False)
            print(f"\n✅ SUCESSO TOTAL! {len(df_final)} títulos salvos em:\n{arquivo_destino_completo}")
        except PermissionError:
            print("\n❌ ERRO DE PERMISSÃO: O arquivo final está aberto no Excel. Feche-o e rode novamente.")
            return False
        except Exception as e:
            print(f"\n❌ ERRO INESPERADO AO SALVAR: {e}")
            return False

    if log_de_erros:
        try:
            df_erros = pd.DataFrame(log_de_erros)
            caminho_log = os.path.join(PASTA_DESTINO, NOME_ARQUIVO_LOG_ERROS)
            df_erros.to_excel(caminho_log, index=False)
            print(f"\n[ATENÇÃO] {len(log_de_erros)} arquivo(s) falharam. Verifique '{NOME_ARQUIVO_LOG_ERROS}'.")
        except Exception as e:
            print(f"Erro ao salvar log de erros: {e}")

    print("\n" + "=" * 50)
    print("RELATÓRIO FINAL DE EXECUÇÃO")
    print("=" * 50)
    print(f"Arquivos processados com sucesso: {len(arquivos_com_sucesso)}")
    print(f"Arquivos com falha:               {len(log_de_erros)}")
    print("=" * 50)
    return True

# ==============================================================================
# EXECUÇÃO / MODO VIGIA
# ==============================================================================

if __name__ == "__main__":
    if not ATIVAR_MODO_VIGIA:
        processar_arquivos_na_pasta()
    else:
        print("--- 🚀 INICIANDO VIGIA DE ARQUIVOS (REVPACK) ---")
        print(f"Aguardando {len(ARQUIVOS_ESPERADOS_VIGIA)} arquivos obrigatórios...")

        while True:
            agora = datetime.datetime.now().strftime('%H:%M:%S')
            ok, faltam = verificar_arquivos_presentes()

            if ok:
                print(f"[{agora}] ✅ TODOS OS ARQUIVOS ENCONTRADOS! INICIANDO...")
                sucesso = processar_arquivos_na_pasta()
                if sucesso:
                    exibir_popup("SUCESSO", "Relatório REVPACK gerado com sucesso!", 64)
                    break
                else:
                    exibir_popup("ERRO", "Ocorreu um erro no processamento. Verifique o console.", 16)
                    break
            else:
                print(f"[{agora}] ⏳ AGUARDANDO. Faltam {len(faltam)} arquivos: {faltam}")
                time.sleep(60)