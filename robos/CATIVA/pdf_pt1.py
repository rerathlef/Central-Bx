import re
import sys
from pathlib import Path
import datetime

try:

    import pdfplumber
except ImportError:
    print("ERRO: pdfplumber não instalado. Execute: pip install pdfplumber")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERRO: openpyxl não instalado. Execute: pip install openpyxl")
    sys.exit(1)
    
# =============================================================================
# 1. CONFIGURAÇÕES
# =============================================================================


if len(sys.argv) > 1 and not sys.argv[1].startswith("-f"):
    PASTA_TRABALHO = sys.argv[1]
else:
    # Se você rodar o arquivo sozinho, ele usa o que estiver aqui:
    PASTA_TRABALHO = r"G:\.shortcut-targets-by-id\1NZDdbjsov2A4Nb33L8NlwmII9RQeNMuu\12 Cativa\03 Financeiro\01. Extratos\2026\05 - Maio\25.05.2026\SICOOB\Nova pasta"




# ---------------------------------------------------------------------------
# Utilitários
# ---------------------------------------------------------------------------

def _get(text: str, campo: str, default: str = "") -> str:
    m = re.search(rf"^{re.escape(campo)}:\s*(.+)", text, re.MULTILINE)
    return m.group(1).strip() if m else default



def _parse_valor(raw) -> float | None:
    if raw is None or raw == "":
        return None
    try:
        # Se já vier como float ou int direto do openpyxl
        if isinstance(raw, (float, int)):
            return float(raw)
            
        # 1. Remove R$, espaços normais e o maldito espaço invisível do banco (\xa0)
        limpo = str(raw).replace("R$", "").replace("\xa0", "").strip()
        # 2. Tira o ponto de milhar
        limpo = limpo.replace(".", "")
        # 3. Troca a vírgula decimal por ponto (padrão que o Python entende)
        limpo = limpo.replace(",", ".")
        # 4. Remove qualquer outra sujeira, mantendo SÓ números, ponto e sinal de menos
        limpo = re.sub(r'[^\d.-]', '', limpo)
        
        return float(limpo)
    except ValueError:
        return None

def _parse_data_pix(raw) -> str:
    if not raw:
        return ""
    # Se o openpyxl já identificar como data no Excel
    if isinstance(raw, datetime.datetime):
        return raw.strftime("%d/%m/%Y")
    
    # Se vier como texto (ex: "02/04/2026 14:30:00")
    texto = str(raw).split(" ")[0].strip()
    return texto


def _parse_data(raw: str) -> str:
    return raw.split(" ")[0].strip() if raw else ""


def _section(text: str, label: str) -> str:
    MARKERS = ["Beneficiário:", "Pagador:", "Beneficiário final:",
               "Datas:", "Valores:", "Situação:", "Autenticação:",
               "Débito:", "Crédito:"]
    m = re.search(rf"^{re.escape(label)}\s*$", text, re.MULTILINE)
    if not m:
        return ""
    start = m.end()
    next_pos = len(text)
    for mk in MARKERS:
        if mk == label:
            continue
        mp = re.search(rf"^{re.escape(mk)}\s*$", text[start:], re.MULTILINE)
        if mp:
            pos = start + mp.start()
            if pos < next_pos:
                next_pos = pos
    return text[start:next_pos]


# ---------------------------------------------------------------------------
# Tipos de registro
# obs: se surgir um tipo novo, vai gerar uma aba "outros"
# ---------------------------------------------------------------------------

TIPO_BOLETO        = "boleto"
TIPO_TRANSFERENCIA = "transferencia"
TIPO_TED           = "ted"
TIPO_CONVENIO      = "convenio"
TIPO_INSS          = "inss"
TIPO_DARF          = "darf"
TIPO_FGTS          = "fgts"
TIPO_DARE          = "dare"
TIPO_OUTRO         = "outro"


def _classifica(text: str) -> str:
    if "EFETIVAÇÃO DE TED" in text:
        return TIPO_TED
    if "COMPROVANTE DE TRANSFERÊNCIA" in text or "ENTRE CONTAS CORRENTES" in text:
        return TIPO_TRANSFERENCIA
    if "PAGAMENTO DE CONVÊNIO" in text:
        return TIPO_CONVENIO
    if "PAGAMENTO INSS" in text or "PAGAMENTO INSS - GPS" in text:
        return TIPO_INSS
    if "PAGAMENTO DARF" in text:
        return TIPO_DARF
    if "RECOLHIMENTO DE FGTS" in text:
        return TIPO_FGTS
    if "DARE-SP/GNRE - SEFAZ-SP" in text:
        return TIPO_DARE
    if "PAGAMENTO DE BOLETO" in text or "PAGAMENTO REJEITADO" in text:
        return TIPO_BOLETO
    return TIPO_OUTRO


# ---------------------------------------------------------------------------
# Parsers
# ---------------------------------------------------------------------------

def _parse_boleto(text: str) -> dict:
    rec = {}
    rec["cooperativa"]          = _get(text, "Cooperativa")
    rec["conta"]                = _get(text, "Conta")
    rec["cliente"]              = _get(text, "Cliente")
    rec["linha_digitavel"]      = _get(text, "Linha digitável")
    rec["numero_documento"]     = _get(text, "Número do documento")
    rec["nosso_numero"]         = _get(text, "Nosso número")
    rec["numero_agendamento"]   = _get(text, "Número do agendamento")
    rec["instituicao_emissora"] = _get(text, "Instituição emissora")
    rec["tipo_documento"]       = _get(text, "Tipo documento")

    benef       = _section(text, "Beneficiário:")
    pagador     = _section(text, "Pagador:")
    benef_final = _section(text, "Beneficiário final:")

    rec["benef_nome"]    = _get(benef, "Nome/Razão Social")
    rec["benef_fantasia"]= _get(benef, "Nome Fantasia")
    rec["benef_cpfcnpj"] = _get(benef, "CPF/CNPJ")

    rec["pagador_nome"]    = _get(pagador, "Nome/Razão social")
    rec["pagador_fantasia"]= _get(pagador, "Nome fantasia")
    rec["pagador_cpfcnpj"] = _get(pagador, "CPF/CNPJ")

    rec["benef_final_nome"]    = _get(benef_final, "Nome/Razão social")
    rec["benef_final_cpfcnpj"] = _get(benef_final, "CPF/CNPJ")

    rec["realizado"]  = _get(text, "Realizado")
    rec["pagamento"]  = _parse_data(_get(text, "Pagamento"))
    rec["vencimento"] = _parse_data(_get(text, "Vencimento"))

    rec["val_documento"] = _parse_valor(_get(text, "Documento"))
    rec["val_desconto"]  = _parse_valor(_get(text, "Desconto/Abatimento"))
    rec["val_juros"]     = _parse_valor(_get(text, "Juros/Multa"))
    rec["val_pago"]      = _parse_valor(_get(text, "Pago"))

    if "PAGAMENTO REJEITADO" in text:
        rec["situacao"] = "Rejeitada/Falha"
    else:
        situacao = _get(text, "Situação")
        if not situacao:
            situacao = "Pago" if rec["pagamento"] else "Agendado (Pendente)"
        rec["situacao"] = situacao

    rec["autenticacao"]= _get(text, "Autenticação")
    return rec


def _parse_transferencia(text: str) -> dict:
    rec = {}
    rec["numero_agendamento"] = _get(text, "Número do agendamento")
    rec["data_agendamento"]   = _get(text, "Data do agendamento")
    rec["data_lancamento"]    = _get(text, "Data do lançamento")
    rec["natureza"]           = _get(text, "Natureza")
    rec["valor"]              = _parse_valor(_get(text, "Valor"))

    deb  = _section(text, "Débito:")
    cred = _section(text, "Crédito:")

    rec["deb_cooperativa"] = _get(deb,  "Cooperativa")
    rec["deb_conta"]       = _get(deb,  "Conta")
    rec["deb_nome"]        = _get(deb,  "Nome") or _get(deb, "Nome/Razão social")
    rec["cred_cooperativa"]= _get(cred, "Cooperativa")
    rec["cred_conta"]      = _get(cred, "Conta")
    rec["cred_nome"]       = _get(cred, "Nome") or _get(cred, "Nome/Razão social")

    rec["autenticacao"] = _get(text, "Autenticação")
    return rec


def _parse_ted(text: str) -> dict:
    rec = {}
    rec["numero_agendamento"] = _get(text, "Número do agendamento")
    rec["data_agendamento"]   = _get(text, "Data do agendamento")
    rec["data_lancamento"]    = _get(text, "Data do lançamento")
    rec["finalidade"]         = _get(text, "Finalidade")
    rec["valor"]              = _parse_valor(_get(text, "Valor"))

    deb  = _section(text, "Débito:")
    cred = _section(text, "Crédito:")

    rec["deb_cooperativa"] = _get(deb,  "Cooperativa")
    rec["deb_conta"]       = _get(deb,  "Conta")
    rec["deb_nome"]        = _get(deb,  "Nome")
    rec["deb_cpfcnpj"]     = _get(deb,  "CPF/CNPJ")

    rec["cred_banco"]      = _get(cred, "Nº Instituição/Banco")
    rec["cred_ispb"]       = _get(cred, "Nº ISPB")
    rec["cred_agencia"]    = _get(cred, "Agência")
    rec["cred_conta"]      = _get(cred, "Conta")
    rec["cred_tipo_conta"] = _get(cred, "Tipo da conta")
    rec["cred_nome"]       = _get(cred, "Nome")
    rec["cred_cpfcnpj"]    = _get(cred, "CPF/CNPJ")

    rec["autenticacao"] = _get(text, "Autenticação")
    return rec


def _parse_convenio(text: str) -> dict:
    rec = {}
    rec["cooperativa"]        = _get(text, "Cooperativa")
    rec["conta"]              = _get(text, "Conta")
    rec["convenio"]           = _get(text, "Convênio")
    rec["codigo_barras"]      = _get(text, "Código de barras")
    rec["numero_agendamento"] = _get(text, "Número do agendamento")
    rec["codigo_nsu"]         = _get(text, "Código NSU")
    rec["data_agendamento"]   = _get(text, "Data do agendamento")
    rec["data_pagamento"]     = _parse_data(_get(text, "Data do pagamento"))
    rec["val_documento"]      = _parse_valor(_get(text, "Valor do documento"))
    rec["val_juros"]          = _parse_valor(_get(text, "Valor dos juros"))
    rec["val_multa"]          = _parse_valor(_get(text, "Valor da multa"))
    rec["val_outros"]         = _parse_valor(_get(text, "Outros encargos"))
    rec["val_desconto"]       = _parse_valor(_get(text, "Valor do desconto"))
    rec["val_total"]          = _parse_valor(_get(text, "Valor total"))
    rec["autenticacao"]       = _get(text, "Autenticação")
    return rec


def _parse_inss(text: str) -> dict:
    rec = {}
    rec["cooperativa"]        = _get(text, "Cooperativa")
    rec["conta"]              = _get(text, "Conta")
    rec["numero_agendamento"] = _get(text, "Número do agendamento")
    rec["data_pagamento"]     = _parse_data(_get(text, "Data do pagamento"))
    rec["codigo_pagamento"]   = _get(text, "Código do pagamento")
    rec["competencia"]        = _get(text, "Competência")
    rec["identificador"]      = _get(text, "Identificador")
    rec["val_inss"]           = _parse_valor(_get(text, "Valor do INSS"))
    rec["val_outras_entidades"] = _parse_valor(_get(text, "Outras entidades"))
    rec["val_multa_juros"]    = _parse_valor(_get(text, "Multa e juros"))
    rec["val_total"]          = _parse_valor(_get(text, "Valor total"))
    rec["autenticacao"]       = _get(text, "Autenticação")
    return rec


def _parse_darf(text: str) -> dict:
    rec = {}
    m_data = re.search(r"DATA DE PAGAMENTO:\s*(\d{2}/\d{2}/\d{4})", text)
    rec["data_pagamento"]     = m_data.group(1) if m_data else ""
    m_valor = re.search(r"VALOR TOTAL:\s*R\$?\s*([0-9.,]+)", text)
    rec["val_total"]          = _parse_valor(m_valor.group(1)) if m_valor else None
    
    rec["banco"]              = _get(text, "BANCO")
    rec["agencia_arrecadador"]= _get(text, "AGÊNCIA ARRECADADOR")
    rec["numero_documento"]   = _get(text, "NÚMERO DO DOCUMENTO")
    rec["numero_agendamento"] = _get(text, "NÚMERO DO AGENDAMENTO")
    rec["codigo_barras"]      = _get(text, "CÓDIGO DE BARRAS")
    rec["periodo_apuracao"]   = _get(text, "PERÍODO DE APURAÇÃO")
    rec["cnpj_cpf"]           = _get(text, "CNPJ / CPF")
    rec["codigo_receita"]     = _get(text, "CÓDIGO DA RECEITA")
    rec["data_vencimento"]    = _parse_data(_get(text, "DATA DO VENCIMENTO"))
    
    m_ci = re.search(r"CI:\s*(.+)", text)
    rec["conta"]              = m_ci.group(1).strip() if m_ci else ""
    rec["autenticacao"]       = _get(text, "AUTENTICAÇÃO")
    return rec


def _parse_fgts(text: str) -> dict:
    rec = {}
    rec["cooperativa"]        = _get(text, "Cooperativa")
    rec["conta"]              = _get(text, "Conta debitada") or _get(text, "Conta")
    rec["codigo_barras"]      = _get(text, "Código de barras")
    rec["cnpj_cei"]           = _get(text, "CNPJ/CEI empresa")
    rec["convenio"]           = _get(text, "Convênio")
    rec["data_validade"]      = _parse_data(_get(text, "Data da validade"))
    rec["competencia"]        = _get(text, "Competência")
    rec["val_recolhido"]      = _parse_valor(_get(text, "Valor recolhido"))
    rec["descricao"]          = _get(text, "Descrição do pagamento")
    rec["codigo_operacao"]    = _get(text, "Código da operação")
    rec["data_pagamento"]     = _parse_data(_get(text, "Data do pagamento"))
    rec["autenticacao"]       = _get(text, "Autenticação")
    return rec


def _parse_dare(text: str) -> dict:
    rec = {}
    rec["cooperativa"]        = _get(text, "Cooperativa")
    rec["conta"]              = _get(text, "Conta")
    m_convenio = re.search(r"Convênio:\s*(.+?)\s", text)
    rec["convenio"]           = m_convenio.group(1) if m_convenio else ""
    m_cnpj = re.search(r"CNPJ:\s*([\d./-]+)", text)
    rec["cnpj"]               = m_cnpj.group(1) if m_cnpj else ""
    m_receita = re.search(r"Receita:\s*(\d+)", text)
    rec["receita"]            = m_receita.group(1) if m_receita else ""
    m_pgto = re.search(r"Data do pagamento:\s*([\d/]+)", text)
    rec["data_pagamento"]     = m_pgto.group(1) if m_pgto else ""
    m_total = re.search(r"Valor total:\s*R\$\s*([\d.,]+)", text)
    rec["val_total"]          = _parse_valor(m_total.group(1)) if m_total else None
    m_controle = re.search(r"Número de controle do DARE:\s*(\d+)", text)
    rec["controle_dare"]      = m_controle.group(1) if m_controle else ""
    rec["autenticacao"]       = _get(text, "Autenticação")
    return rec


def _parse_outro(text: str) -> dict:
    # Tenta extrair o máximo possível de forma genérica
    tipo_match = re.search(r"COMPROVANTE DE\n(.+)", text)
    tipo_comp  = tipo_match.group(1).strip() if tipo_match else ""
    return {
        "tipo_comprovante":   tipo_comp,
        "numero_agendamento": _get(text, "Número do agendamento"),
        "cooperativa":        _get(text, "Cooperativa"),
        "conta":              _get(text, "Conta"),
        "data_pagamento":     _parse_data(_get(text, "Data do pagamento")),
        "valor":              _parse_valor(_get(text, "Valor")),
        "autenticacao":       _get(text, "Autenticação"),
        "texto_resumo":       text[:600].replace("\n", " | "),
    }


def extract_from_pix_excel(xlsx_path: str) -> list[dict]:
    recs = []
    
    print(f"  Lendo extrato PIX: {Path(xlsx_path).name}")
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    
    header_row_idx = None
    col_map = {}
    
    # 1. Encontra a linha de cabeçalho dinamicamente
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
        if row and "Valor" in row and "Destinatário" in row:
            header_row_idx = row_idx
            for col_idx, cell_value in enumerate(row):
                
                # AQUI ESTÁ A MÁGICA: Aceita as duas formas e salva como "Data_PIX"
                if cell_value in ["Data/Hora Movimento", "Data Movimento"]:
                    col_map["Data_PIX"] = col_idx
                    
                elif cell_value == "Destinatário":
                    col_map["Destinatário"] = col_idx
                    
                elif cell_value == "Valor":
                    col_map["Valor"] = col_idx
            break
            
    if not header_row_idx:
        print(f"    ERRO: Cabeçalhos não encontrados no arquivo {xlsx_path}")
        return []

    # 2. Lê as linhas abaixo do cabeçalho
    for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
        
        # Pega as posições das colunas com segurança (se não achar, retorna None)
        idx_valor = col_map.get("Valor")
        idx_data = col_map.get("Data_PIX")
        idx_dest = col_map.get("Destinatário")
        
        # Proteção: se a linha for vazia, se a coluna valor não existir, ou se a célula estiver vazia, pula
        if not row or idx_valor is None or idx_valor >= len(row) or row[idx_valor] is None:
            continue
            
        # Puxa os valores usando os índices seguros
        raw_data = row[idx_data] if idx_data is not None and idx_data < len(row) else ""
        raw_dest = row[idx_dest] if idx_dest is not None and idx_dest < len(row) else ""
        raw_valor = row[idx_valor]

        data_limpa = _parse_data_pix(raw_data)
        valor_limpo = _parse_valor(raw_valor)

        if valor_limpo is None:
            continue

        rec = {
            "_tipo_operacao": "PIX",
            "_vg_situacao": "Efetivada",
            "_vg_data": data_limpa,
            "_vg_pagador": "",
            "_vg_conta_origem": "",
            "_vg_beneficiario": str(raw_dest).strip() if raw_dest else "",
            "_vg_conta_destino": "",
            "_vg_valor": valor_limpo,
        }
        recs.append(rec)
        
    return recs


# ---------------------------------------------------------------------------
# Extração do PDF
# ---------------------------------------------------------------------------

def extract_from_pdf(pdf_path: str) -> dict[str, list[dict]]:
    result: dict[str, list[dict]] = {
        TIPO_BOLETO:          [],
        TIPO_TRANSFERENCIA:   [],
        TIPO_TED:             [],
        TIPO_CONVENIO:        [],
        TIPO_INSS:            [],
        TIPO_DARF:            [],
        TIPO_FGTS:            [],
        TIPO_DARE:            [],
        TIPO_OUTRO:           [],
    }

    path = Path(pdf_path)
    if not path.exists():
        print(f"ERRO: arquivo '{pdf_path}' não encontrado.")
        return result

    print(f"\nProcessando: {path.name}")

    PARSERS = {
        TIPO_BOLETO:          _parse_boleto,
        TIPO_TRANSFERENCIA:   _parse_transferencia,
        TIPO_TED:             _parse_ted,
        TIPO_CONVENIO:        _parse_convenio,
        TIPO_INSS:            _parse_inss,
        TIPO_DARF:            _parse_darf,
        TIPO_FGTS:            _parse_fgts,
        TIPO_DARE:            _parse_dare,
        TIPO_OUTRO:           _parse_outro,
    }

    with pdfplumber.open(str(path)) as pdf:
        total = len(pdf.pages)
        print(f"  Total de páginas: {total}")

        for i, page in enumerate(pdf.pages, 1):
            text = page.extract_text() or ""
            if not text.strip():
                continue

            tipo = _classifica(text)
            rec  = PARSERS[tipo](text)
            rec["_pagina"]  = i
            rec["_arquivo"] = path.name
            
            # --- Adicionando campos genéricos para a Visão Geral ---
            rec["_tipo_operacao"] = tipo
            if tipo == TIPO_BOLETO:
                rec["_vg_pagador"] = rec.get("pagador_nome") or rec.get("cliente")
                rec["_vg_beneficiario"] = rec.get("benef_nome")
                rec["_vg_conta_origem"] = rec.get("conta")
                rec["_vg_conta_destino"] = ""  # boleto não tem conta de destino clara
                rec["_vg_valor"] = rec.get("val_pago") or rec.get("val_documento")
                rec["_vg_data"] = rec.get("pagamento") or rec.get("vencimento") or rec.get("realizado", "")[:10]
                rec["_vg_situacao"] = rec.get("situacao")
            elif tipo == TIPO_TRANSFERENCIA:
                rec["_vg_pagador"] = rec.get("deb_nome")
                rec["_vg_beneficiario"] = rec.get("cred_nome")
                rec["_vg_conta_origem"] = rec.get("deb_conta")
                rec["_vg_conta_destino"] = rec.get("cred_conta")
                rec["_vg_valor"] = rec.get("valor")
                rec["_vg_data"] = rec.get("data_lancamento") or rec.get("data_agendamento")
                rec["_vg_situacao"] = "Efetivada"
            elif tipo == TIPO_TED:
                rec["_vg_pagador"] = rec.get("deb_nome")
                rec["_vg_beneficiario"] = rec.get("cred_nome")
                rec["_vg_conta_origem"] = rec.get("deb_conta")
                rec["_vg_conta_destino"] = f"{rec.get('cred_banco', '')} - {rec.get('cred_conta', '')}".strip(" -")
                rec["_vg_valor"] = rec.get("valor")
                rec["_vg_data"] = rec.get("data_lancamento") or rec.get("data_agendamento")
                rec["_vg_situacao"] = "Efetivada"
            elif tipo == TIPO_CONVENIO:
                rec["_vg_pagador"] = ""  # Não especificado claramente no comprovante de convênio
                rec["_vg_beneficiario"] = rec.get("convenio")
                rec["_vg_conta_origem"] = rec.get("conta")
                rec["_vg_conta_destino"] = ""
                rec["_vg_valor"] = rec.get("val_total") or rec.get("val_documento")
                rec["_vg_data"] = rec.get("data_pagamento") or rec.get("data_agendamento")[:10]
                rec["_vg_situacao"] = "Efetivada"
            elif tipo == TIPO_INSS:
                rec["_vg_pagador"] = ""
                rec["_vg_beneficiario"] = "INSS/GPS (" + str(rec.get("codigo_pagamento","")) + ")"
                rec["_vg_conta_origem"] = rec.get("conta")
                rec["_vg_conta_destino"] = ""
                rec["_vg_valor"] = rec.get("val_total")
                rec["_vg_data"] = rec.get("data_pagamento")
                rec["_vg_situacao"] = "Efetivada"
            elif tipo == TIPO_DARF:
                rec["_vg_pagador"] = ""
                rec["_vg_beneficiario"] = "Receita Federal (DARF " + str(rec.get("codigo_receita","")) + ")"
                rec["_vg_conta_origem"] = rec.get("conta")
                rec["_vg_conta_destino"] = ""
                rec["_vg_valor"] = rec.get("val_total")
                rec["_vg_data"] = rec.get("data_pagamento")
                rec["_vg_situacao"] = "Efetivada"
            elif tipo == TIPO_FGTS:
                rec["_vg_pagador"] = ""
                rec["_vg_beneficiario"] = "Caixa Econômica (FGTS)"
                rec["_vg_conta_origem"] = rec.get("conta")
                rec["_vg_conta_destino"] = ""
                rec["_vg_valor"] = rec.get("val_recolhido")
                rec["_vg_data"] = rec.get("data_pagamento")
                rec["_vg_situacao"] = "Efetivada"
            elif tipo == TIPO_DARE:
                rec["_vg_pagador"] = ""
                rec["_vg_beneficiario"] = "SEFAZ-SP (DARE-SP/GNRE)"
                rec["_vg_conta_origem"] = rec.get("conta")
                rec["_vg_conta_destino"] = ""
                rec["_vg_valor"] = rec.get("val_total")
                rec["_vg_data"] = rec.get("data_pagamento")
                rec["_vg_situacao"] = "Efetivada"
            else:
                rec["_vg_pagador"] = ""
                rec["_vg_beneficiario"] = "Outro"
                rec["_vg_conta_origem"] = rec.get("conta", "")
                rec["_vg_conta_destino"] = ""
                rec["_vg_valor"] = rec.get("valor")
                rec["_vg_data"] = rec.get("data_pagamento")
                rec["_vg_situacao"] = "Revisar"

            result[tipo].append(rec)

            if i % 50 == 0:
                print(f"  ... {i}/{total} páginas processadas")

    # Resumo no console
    labels = {
        TIPO_BOLETO:          "Boletos (Pagos/Agendados/Rejeitados)",
        TIPO_TRANSFERENCIA:   "Transferências",
        TIPO_TED:             "TEDs",
        TIPO_CONVENIO:        "Convênios",
        TIPO_INSS:            "INSS/GPS",
        TIPO_DARF:            "DARF",
        TIPO_FGTS:            "FGTS",
        TIPO_DARE:            "DARE-SP",
        TIPO_OUTRO:           "Outros",
    }
    for tipo, label in labels.items():
        n = len(result[tipo])
        if n:
            print(f"  ✓ {label}: {n}")
    return result


# ---------------------------------------------------------------------------
# Definição de colunas por tipo
# ---------------------------------------------------------------------------
# Cada item é (label, record_key, width)
# As colunas de debug (_pagina, _arquivo) são adicionadas dinamicamente.

DEBUG_COLS = [
    ("Arquivo", "_arquivo", 38),
    ("Página",  "_pagina",   7),
]

COLS_BOLETO = [
    ("Pagamento",             "pagamento",          14),
    ("Vencimento",            "vencimento",         14),
    ("Realizado",             "realizado",          22),
    ("Situação",              "situacao",           12),
    ("Conta",                 "conta",              12),
    ("Cliente",               "cliente",            35),
    ("Tipo Doc.",             "tipo_documento",     14),
    ("Nº Documento",          "numero_documento",   18),
    ("Nº Agendamento",        "numero_agendamento", 18),
    ("Nosso Número",          "nosso_numero",       20),
    ("Beneficiário Nome",     "benef_nome",         35),
    ("Beneficiário Fantasia", "benef_fantasia",     35),
    ("Beneficiário CNPJ",     "benef_cpfcnpj",      22),
    ("Pagador Nome",          "pagador_nome",       35),
    ("Pagador Fantasia",      "pagador_fantasia",   35),
    ("Pagador CNPJ",          "pagador_cpfcnpj",    22),
    ("Benef. Final Nome",     "benef_final_nome",   35),
    ("Benef. Final CNPJ",     "benef_final_cpfcnpj",22),
    ("Valor Documento",       "val_documento",      16),
    ("Desconto/Abatimento",   "val_desconto",       18),
    ("Juros/Multa",           "val_juros",          14),
    ("Valor Pago",            "val_pago",           14),
    ("Cooperativa",           "cooperativa",        40),
    ("Inst. Emissora",        "instituicao_emissora",18),
    ("Linha Digitável",       "linha_digitavel",    55),
    ("Autenticação",          "autenticacao",       40),
]
MONETARY_BOLETO = {"val_documento", "val_desconto", "val_juros", "val_pago"}

COLS_TRANSFERENCIA = [
    ("Data Agendamento",  "data_agendamento",   16),
    ("Data Lançamento",   "data_lancamento",    16),
    ("Nº Agendamento",    "numero_agendamento", 18),
    ("Natureza",          "natureza",           40),
    ("Valor",             "valor",              14),
    ("Déb. Cooperativa",  "deb_cooperativa",    30),
    ("Déb. Conta",        "deb_conta",          14),
    ("Déb. Nome",         "deb_nome",           35),
    ("Cred. Cooperativa", "cred_cooperativa",   30),
    ("Cred. Conta",       "cred_conta",         14),
    ("Cred. Nome",        "cred_nome",          35),
    ("Autenticação",      "autenticacao",       40),
]
MONETARY_TRANSF = {"valor"}

COLS_TED = [
    ("Data Agendamento",  "data_agendamento",   16),
    ("Data Lançamento",   "data_lancamento",    16),
    ("Nº Agendamento",    "numero_agendamento", 18),
    ("Finalidade",        "finalidade",         35),
    ("Valor",             "valor",              14),
    ("Déb. Cooperativa",  "deb_cooperativa",    30),
    ("Déb. Conta",        "deb_conta",          14),
    ("Déb. Nome",         "deb_nome",           35),
    ("Déb. CNPJ",         "deb_cpfcnpj",        22),
    ("Cred. Banco",       "cred_banco",         35),
    ("Cred. ISPB",        "cred_ispb",          14),
    ("Cred. Agência",     "cred_agencia",       14),
    ("Cred. Conta",       "cred_conta",         14),
    ("Cred. Tipo Conta",  "cred_tipo_conta",    18),
    ("Cred. Nome",        "cred_nome",          35),
    ("Cred. CNPJ",        "cred_cpfcnpj",       22),
    ("Autenticação",      "autenticacao",       40),
]
MONETARY_TED = {"valor"}

COLS_CONVENIO = [
    ("Data Pagamento",    "data_pagamento",     16),
    ("Data Agendamento",  "data_agendamento",   20),
    ("Nº Agendamento",    "numero_agendamento", 18),
    ("Convênio",          "convenio",           30),
    ("Conta",             "conta",              14),
    ("Cooperativa",       "cooperativa",        30),
    ("Código NSU",        "codigo_nsu",         20),
    ("Código de Barras",  "codigo_barras",      55),
    ("Valor Documento",   "val_documento",      16),
    ("Juros",             "val_juros",          12),
    ("Multa",             "val_multa",          12),
    ("Outros Encargos",   "val_outros",         14),
    ("Desconto",          "val_desconto",       12),
    ("Valor Total",       "val_total",          14),
    ("Autenticação",      "autenticacao",       40),
]
MONETARY_CONVENIO = {"val_documento","val_juros","val_multa","val_outros","val_desconto","val_total"}

COLS_INSS = [
    ("Data Pagamento",      "data_pagamento",       16),
    ("Nº Agendamento",      "numero_agendamento",   18),
    ("Cooperativa",         "cooperativa",          30),
    ("Conta",               "conta",                14),
    ("Código Pagamento",    "codigo_pagamento",     18),
    ("Competência",         "competencia",          14),
    ("Identificador",       "identificador",        22),
    ("Valor INSS",          "val_inss",             14),
    ("Outras Entidades",    "val_outras_entidades", 18),
    ("Multa e Juros",       "val_multa_juros",      14),
    ("Valor Total",         "val_total",            14),
    ("Autenticação",        "autenticacao",         40),
]
MONETARY_INSS = {"val_inss","val_outras_entidades","val_multa_juros","val_total"}

COLS_DARF = [
    ("Data Pagamento",    "data_pagamento",      16),
    ("Conta / Origem",    "conta",               40),
    ("Banco Arrecadador", "banco",               40),
    ("Agência Arrecadador","agencia_arrecadador",30),
    ("Nº Documento",      "numero_documento",    20),
    ("Nº Agendamento",    "numero_agendamento",  18),
    ("CNPJ / CPF",        "cnpj_cpf",            20),
    ("Cód. Receita",      "codigo_receita",      14),
    ("Período Apuração",  "periodo_apuracao",    20),
    ("Data Vencimento",   "data_vencimento",     16),
    ("Código Barras",     "codigo_barras",       55),
    ("Valor Total",       "val_total",           16),
    ("Autenticação",      "autenticacao",        40),
]
MONETARY_DARF = {"val_total"}

COLS_FGTS = [
    ("Data Pagamento",    "data_pagamento",      16),
    ("Cooperativa",       "cooperativa",         30),
    ("Conta",             "conta",               40),
    ("CNPJ / CEI",        "cnpj_cei",            20),
    ("Competência",       "competencia",         14),
    ("Data Validade",     "data_validade",       16),
    ("Convênio",          "convenio",            14),
    ("Código Operação",   "codigo_operacao",     18),
    ("Código Barras",     "codigo_barras",       55),
    ("Descrição",         "descricao",           40),
    ("Valor Recolhido",   "val_recolhido",       16),
    ("Autenticação",      "autenticacao",        40),
]
MONETARY_FGTS = {"val_recolhido"}

COLS_DARE = [
    ("Data Pagamento",    "data_pagamento",      16),
    ("Cooperativa",       "cooperativa",         30),
    ("Conta",             "conta",               40),
    ("Convênio",          "convenio",            30),
    ("CNPJ",              "cnpj",                20),
    ("Receita",           "receita",             14),
    ("Nº Controle DARE",  "controle_dare",       25),
    ("Valor Total",       "val_total",           16),
    ("Autenticação",      "autenticacao",        40),
]
MONETARY_DARE = {"val_total"}

COLS_OUTRO = [
    ("Tipo Comprovante",  "tipo_comprovante",     30),
    ("Data Pagamento",    "data_pagamento",        16),
    ("Nº Agendamento",    "numero_agendamento",    18),
    ("Cooperativa",       "cooperativa",           30),
    ("Conta",             "conta",                 14),
    ("Valor",             "valor",                 14),
    ("Autenticação",      "autenticacao",          40),
    ("Texto (resumo)",    "texto_resumo",          80),  # mantido só aqui pois é genérico
]
MONETARY_OUTRO = {"valor"}

COLS_VISAO_GERAL = [
    ("Tipo Operação",     "_tipo_operacao",       16),
    ("Situação",          "_vg_situacao",         20),
    ("Data",              "_vg_data",             14),
    ("Pagador",           "_vg_pagador",          40),
    ("Conta Origem",      "_vg_conta_origem",     18),
    ("Beneficiário",      "_vg_beneficiario",     40),
    ("Conta Destino",     "_vg_conta_destino",    22),
    ("Valor",             "_vg_valor",            16),
]
MONETARY_VG = {"_vg_valor"}


# ---------------------------------------------------------------------------
# Helpers de estilo / escrita
# ---------------------------------------------------------------------------

def _style_header(ws, fgColor: str = "1F3864"):
    fill  = PatternFill("solid", fgColor=fgColor)
    font  = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin  = Side(style="thin", color="FFFFFF")
    bord  = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[1]:
        cell.fill = fill; cell.font = font
        cell.alignment = align; cell.border = bord
    ws.row_dimensions[1].height = 28


def _zebra(ws, row_num: int, even_color: str = "EBF1F9"):
    color = even_color if row_num % 2 == 0 else "FFFFFF"
    fill  = PatternFill("solid", fgColor=color)
    font  = Font(name="Calibri", size=9)
    align = Alignment(vertical="center")
    for cell in ws[row_num]:
        cell.fill = fill; cell.font = font; cell.alignment = align


def _write_sheet(ws, records: list[dict], cols: list[tuple],
                 monetary_keys: set, header_color: str, even_color: str,
                 include_debug: bool = False):
    """
    Escreve registros na aba.
    Se include_debug=True, prefixa as colunas DEBUG_COLS (Arquivo, Página).
    """
    final_cols = (DEBUG_COLS + cols) if include_debug else cols
    ws.append([c[0] for c in final_cols])
    _style_header(ws, header_color)

    valor_fmt = '#,##0.00'
    for row_idx, rec in enumerate(records, 2):
        ws.append([rec.get(c[1], "") for c in final_cols])
        for col_idx, (_, key, _) in enumerate(final_cols, 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if key in monetary_keys:
                cell.number_format = valor_fmt
                cell.alignment = Alignment(horizontal="right", vertical="center")
        _zebra(ws, row_idx, even_color)

    for col_idx, (_, __, width) in enumerate(final_cols, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def _build_visao_geral(ws, all_records: list[dict], include_debug: bool):
    """Constrói a aba unificada e ordenada por data."""
    records_sorted = sorted(all_records, key=lambda r: r.get("_vg_data", ""))
    _write_sheet(ws, records_sorted, COLS_VISAO_GERAL, MONETARY_VG,
                 header_color="2B2B2B", even_color="F0F0F0",
                 include_debug=include_debug)


# ---------------------------------------------------------------------------
# Geração do Excel (uma função reutilizável)
# ---------------------------------------------------------------------------

SHEET_DEFS = [
    # (tipo,               nome_aba,           cols,               monetary,           header_color, even_color, sort_key)
    (TIPO_BOLETO,          "Boletos",           COLS_BOLETO,        MONETARY_BOLETO,    "1F3864",     "EBF1F9",   lambda r: (r.get("pagamento","") or r.get("vencimento",""), r.get("realizado",""))),
    (TIPO_TRANSFERENCIA,   "Transferências",    COLS_TRANSFERENCIA, MONETARY_TRANSF,    "1A4D2E",     "E8F5E9",   lambda r:  r.get("data_lancamento","")),
    (TIPO_TED,             "TEDs",              COLS_TED,           MONETARY_TED,       "1A3A5C",     "E3EFF8",   lambda r:  r.get("data_lancamento","")),
    (TIPO_CONVENIO,        "Convênios",         COLS_CONVENIO,      MONETARY_CONVENIO,  "5C3A1A",     "FFF3E0",   lambda r:  r.get("data_pagamento","")),
    (TIPO_INSS,            "INSS-GPS",          COLS_INSS,          MONETARY_INSS,      "5C1A3A",     "FDE8F0",   lambda r:  r.get("data_pagamento","")),
    (TIPO_DARF,            "DARF",              COLS_DARF,          MONETARY_DARF,      "D97706",     "FEF3C7",   lambda r:  r.get("data_pagamento","")),
    (TIPO_FGTS,            "FGTS",              COLS_FGTS,          MONETARY_FGTS,      "047857",     "D1FAE5",   lambda r:  r.get("data_pagamento","")),
    (TIPO_DARE,            "DARE-SP",           COLS_DARE,          MONETARY_DARE,      "9333EA",     "F3E8FF",   lambda r:  r.get("data_pagamento","")),
    (TIPO_OUTRO,           "Outros",            COLS_OUTRO,         MONETARY_OUTRO,     "555555",     "F5F5F5",   lambda r:  r.get("data_pagamento","")),
]


def _build_workbook(result: dict, include_debug: bool) -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Coleta TODOS os registros para a Visão Geral
    all_recs = []
    for tipo, recs in result.items():
        all_recs.extend(recs)

    # 1) Aba Visão Geral (a primeira de todas)
    if all_recs:
        ws_vg = wb.create_sheet("Visão Geral")
        _build_visao_geral(ws_vg, all_recs, include_debug)

    # 2) Abas específicas
    for tipo, nome, cols, monetary, hcolor, ecolor, sort_fn in SHEET_DEFS:
        recs = result[tipo]
        if not recs:
            continue
        recs_sorted = sorted(recs, key=sort_fn)
        ws = wb.create_sheet(nome)
        _write_sheet(ws, recs_sorted, cols, monetary, hcolor, ecolor,
                     include_debug=include_debug)

    return wb


def generate_excel(result: dict[str, list[dict]], base_path: str):
    """Gera dois arquivos: _OFICIAL.xlsx e _DEBUG.xlsx."""
    base = Path(base_path)
    stem = base.stem.replace("_OFICIAL", "").replace("_DEBUG", "")
    parent = base.parent

    path_oficial = parent / f"{stem}.xlsx"
   # path_debug   = parent / f"{stem}_DEBUG.xlsx"

    wb_oficial = _build_workbook(result, include_debug=False)
    wb_oficial.save(str(path_oficial))
    print(f"\n✅ OFICIAL: {path_oficial.name}")

    wb_debug = _build_workbook(result, include_debug=True)
    #wb_debug.save(str(path_debug))
    #print(f"✅ DEBUG:   {path_debug.name}")

    # Contagens
    totais = {k: len(v) for k, v in result.items()}
    total  = sum(totais.values())
    print(f"\n   Boletos:            {totais.get(TIPO_BOLETO,0)}")
    print(f"   Transferências:     {totais.get(TIPO_TRANSFERENCIA,0)}")
    print(f"   TEDs:               {totais.get(TIPO_TED,0)}")
    print(f"   Convênios:          {totais.get(TIPO_CONVENIO,0)}")
    print(f"   INSS / GPS:         {totais.get(TIPO_INSS,0)}")
    print(f"   DARF:               {totais.get(TIPO_DARF,0)}")
    print(f"   FGTS:               {totais.get(TIPO_FGTS,0)}")
    print(f"   DARE-SP:            {totais.get(TIPO_DARE,0)}")
    print(f"   Outros:             {totais.get(TIPO_OUTRO,0)}")
    print(f"   TOTAL:              {total}")


# ---------------------------------------------------------------------------
# Ponto de entrada
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Ponto de entrada
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Ponto de entrada
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Ponto de entrada
# ---------------------------------------------------------------------------

def main():
    
    pasta = Path(PASTA_TRABALHO)
    
    if not pasta.exists() or not pasta.is_dir():
        print(f"ERRO: A pasta '{PASTA_TRABALHO}' não foi encontrada.")
        sys.exit(1)

    pdf_files = list(pasta.glob("*.[pP][dD][fF]"))

    print(f"Encontrados {len(pdf_files)} arquivo(s) PDF. Iniciando processamento...\n")
    
    # Criamos um "caderninho" para anotar os PIXs que já foram casados com PDFs
    pix_processados = set()

    for pdf_file in pdf_files:
        print("-" * 50)
        
        # Cria o dicionário zerado para este arquivo
        result_individual: dict[str, list[dict]] = {t: [] for t in [
            TIPO_BOLETO, TIPO_TRANSFERENCIA,
            TIPO_TED, TIPO_CONVENIO, TIPO_INSS, TIPO_DARF, TIPO_FGTS, TIPO_DARE, TIPO_OUTRO
        ]}

        # 1. Extrai os dados do PDF
        result = extract_from_pdf(str(pdf_file))
        for tipo, recs in result.items():
            result_individual[tipo].extend(recs)

        # 2. Descobre o nome da empresa para buscar o arquivo PIX
        # Ex: "COMPROVANTES - CATIVA APIUNA" -> "CATIVA APIUNA"
        nome_base = pdf_file.stem.replace("COMPROVANTES - ", "").strip()
        
        # Procura arquivos xlsx que tenham o nome da empresa E a palavra PIX
        # Ignora arquivos temporários do Excel que começam com "~$"
        pix_files = [f for f in pasta.glob(f"*{nome_base}*.xlsx") 
                     if "PIX" in f.name.upper() and not f.name.startswith("~$")]

        # 3. Se achar o Excel do PIX correspondente, faz a leitura!
        if pix_files:
            arquivo_pix = pix_files[0] # Pega o primeiro que bater com o critério
            recs_pix = extract_from_pix_excel(str(arquivo_pix))
            
            # Jogamos na categoria OUTRO, mas eles vão cruzar certinho na aba VISÃO GERAL
            result_individual[TIPO_OUTRO].extend(recs_pix)
            
            # ANOTA NO CADERNO QUE ESSE PIX JÁ FOI LIDO
            pix_processados.add(arquivo_pix) 
            
            print(f"  ✓ PIX processado: {len(recs_pix)} registro(s) adicionado(s).")
        else:
            print(f"  ⚠ Nenhum arquivo de PIX encontrado para: {nome_base}")

        total = sum(len(v) for v in result_individual.values())
        if total == 0:
            print(f"\nNenhum registro encontrado para {pdf_file.name}. Pulando...")
            continue

        out_base = str(pdf_file)

        # Gera o Excel final com o PDF + PIX juntos!
        generate_excel(result_individual, out_base)
        
    # =========================================================================
    # FASE 2: RESGATAR OS PIXs "ÓRFÃOS" (QUE NÃO TINHAM PDF)
    # =========================================================================
    todos_pix = [f for f in pasta.glob("*.xlsx") if "PIX" in f.name.upper() and not f.name.startswith("~$")]
    
    for arquivo_pix in todos_pix:
        # Se o arquivo não está no "caderninho", significa que ele sobrou
        if arquivo_pix not in pix_processados:
            print("-" * 50)
            print(f"Processando PIX avulso (sem PDF): {arquivo_pix.name}")
            
            result_individual = {t: [] for t in [TIPO_BOLETO, TIPO_TRANSFERENCIA, TIPO_TED, TIPO_CONVENIO, TIPO_INSS, TIPO_DARF, TIPO_FGTS, TIPO_DARE, TIPO_OUTRO]}
            
            recs_pix = extract_from_pix_excel(str(arquivo_pix))
            if recs_pix:
                result_individual[TIPO_OUTRO].extend(recs_pix)
                
                # Transforma "COMPROVANTES PAGAMENTOS PIX SICOOB - CDT" em "COMPROVANTES - CDT"
                nome_limpo = arquivo_pix.stem.replace("COMPROVANTES PAGAMENTOS PIX SICOOB", "COMPROVANTES")
                
                # Passa com um ".pdf" fantasma só pro seu gerador de Excel trocar pra ".xlsx" certinho
                out_base = str(pasta / f"{nome_limpo}.pdf") 
                
                # Gera o Excel final com o PDF + PIX juntos!
        generate_excel(result_individual, out_base)
        
    # =========================================================================
    # FASE 2: RESGATAR OS PIXs "ÓRFÃOS" (QUE NÃO TINHAM PDF)
    # =========================================================================
    todos_pix = [f for f in pasta.glob("*.xlsx") if "PIX" in f.name.upper() and not f.name.startswith("~$")]
    
    for arquivo_pix in todos_pix:
        # Se o arquivo não está no "caderninho", significa que ele sobrou
        if arquivo_pix not in pix_processados:
            print("-" * 50)
            print(f"Processando PIX avulso (sem PDF): {arquivo_pix.name}")
            
            result_individual = {t: [] for t in [TIPO_BOLETO, TIPO_TRANSFERENCIA, TIPO_TED, TIPO_CONVENIO, TIPO_INSS, TIPO_DARF, TIPO_FGTS, TIPO_DARE, TIPO_OUTRO]}
            
            recs_pix = extract_from_pix_excel(str(arquivo_pix))
            if recs_pix:
                result_individual[TIPO_OUTRO].extend(recs_pix)
                
                # Transforma "COMPROVANTES PAGAMENTOS PIX SICOOB - CDT" em "COMPROVANTES - CDT"
                nome_limpo = arquivo_pix.stem.replace("COMPROVANTES PAGAMENTOS PIX SICOOB", "COMPROVANTES")
                
                # Passa com um ".pdf" fantasma só pro seu gerador de Excel trocar pra ".xlsx" certinho
                out_base = str(pasta / f"{nome_limpo}.pdf") 
                
                generate_excel(result_individual, out_base)
            else:
                print("  Nenhum registro encontrado neste PIX.")

    print(f"\n{'-' * 50}")
    print("Concluído! O robô finalizou todos os arquivos (PDFs e PIXs).")
if __name__ == "__main__":
    main()
    
    
    