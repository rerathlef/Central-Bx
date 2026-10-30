"""
drive_utils.py
=====================================================================
Módulo unificado de conexão com Google Drive via OAuth 2.0 (Central Bx Finance).
Gerencia a navegação de pastas, download para diretório temporário isolado
e upload dos relatórios finais para múltiplos projetos (Cativa, Revpack, etc.).
=====================================================================
"""

import io
import os
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from google.oauth2.credentials import Credentials

SCOPES = ["https://www.googleapis.com/auth/drive"]

# =====================================================================
# IDs DAS PASTAS RAIZ NO GOOGLE DRIVE POR PROJETO
# =====================================================================
ID_PASTA_RAIZ_CATIVA = "1F6NiwgME7xPe_xUQQwzQHwwaBenRXO6N"
ID_PASTA_REVPACK_CONCILIACAO = "1a5CLnMLhinJAAqdD9Tdr3fvFBtcf_hj0"
ID_PASTA_REVPACK_SACADOS = "1a5CLnMLhinJAAqdD9Tdr3fvFBtcf_hj0"

# Compatibilidade com chamadas padrão
ID_PASTA_RAIZ = ID_PASTA_RAIZ_CATIVA


def conectar_drive(credenciais_dict):
    """
    Conecta na API do Google Drive usando os dados OAuth 2.0 salvos nos Secrets.
    Remove 'expiry' para forçar renovação automática via refresh_token.
    """
    info = dict(credenciais_dict)

    if "expiry" in info:
        del info["expiry"]

    creds = Credentials.from_authorized_user_info(info, SCOPES)
    servico = build("drive", "v3", credentials=creds)

    return servico


def _achar_subpasta(servico, nome_pasta, id_pasta_pai):
    """Procura uma subpasta pelo nome dentro de uma pasta pai."""
    nome_escaped = nome_pasta.replace("'", "\\'")
    query = (
        f"'{id_pasta_pai}' in parents "
        f"and name = '{nome_escaped}' "
        f"and mimeType = 'application/vnd.google-apps.folder' "
        f"and trashed = false"
    )
    resultado = servico.files().list(
        q=query, fields="files(id, name)", pageSize=10
    ).execute()
    arquivos = resultado.get("files", [])
    if arquivos:
        return arquivos[0]["id"]

    # Fallback: busca ignorando diferenças de maiúsculas/minúsculas ou espaços nas pontas
    query_all = (
        f"'{id_pasta_pai}' in parents "
        f"and mimeType = 'application/vnd.google-apps.folder' "
        f"and trashed = false"
    )
    res_all = servico.files().list(
        q=query_all, fields="files(id, name)", pageSize=100
    ).execute()
    for arq in res_all.get("files", []):
        if arq["name"].strip().upper() == nome_pasta.strip().upper():
            return arq["id"]

    raise FileNotFoundError(
        f"Pasta '{nome_pasta}' não encontrada dentro da pasta pai (id={id_pasta_pai})."
    )


def navegar_caminho(servico, caminho_relativo, id_raiz=ID_PASTA_RAIZ):
    """
    Navega pasta por pasta a partir de `id_raiz` e devolve o ID da pasta final.
    Aceita tanto barra normal '/' quanto contrabarra '\\'.
    Se o usuário incluir o próprio nome da pasta raiz no início do caminho,
    detecta e pula automaticamente para evitar erro de subpasta.
    """
    caminho_normalizado = caminho_relativo.replace("\\", "/")
    partes = [p.strip() for p in caminho_normalizado.split("/") if p.strip()]

    id_atual = id_raiz

    nomes_raiz_conhecidos = {
        "12 CATIVA",
        "08 REVPACK",
        "EXTRATOS BANCÁRIOS",
        "EXTRATOS BANCARIOS",
        "POSIÇÕES FUNDOS",
        "POSICOES FUNDOS",
    }
    try:
        meta_raiz = servico.files().get(fileId=id_raiz, fields="name").execute()
        nome_raiz_real = meta_raiz.get("name", "").strip().upper()
        if nome_raiz_real:
            nomes_raiz_conhecidos.add(nome_raiz_real)
    except Exception:
        pass

    for idx, nome_parte in enumerate(partes):
        if idx == 0 and id_atual == id_raiz:
            if nome_parte.strip().upper() == "12 CATIVA":
                continue
            try:
                id_atual = _achar_subpasta(servico, nome_parte, id_atual)
                continue
            except FileNotFoundError:
                if nome_parte.strip().upper() in nomes_raiz_conhecidos:
                    continue
                raise

        id_atual = _achar_subpasta(servico, nome_parte, id_atual)

    return id_atual


def listar_arquivos_da_pasta(servico, id_pasta):
    """Lista todos os arquivos dentro de uma pasta do Drive."""
    arquivos = []
    page_token = None
    while True:
        query = f"'{id_pasta}' in parents and trashed = false"
        resposta = servico.files().list(
            q=query,
            fields="nextPageToken, files(id, name, mimeType)",
            pageToken=page_token,
            pageSize=100,
        ).execute()
        arquivos.extend(resposta.get("files", []))
        page_token = resposta.get("nextPageToken")
        if not page_token:
            break
    return arquivos


def baixar_pasta_inteira(servico, id_pasta, destino_local):
    """
    Baixa todos os arquivos de uma pasta do Drive para uma pasta temporária local.
    Ignora subpastas.
    """
    os.makedirs(destino_local, exist_ok=True)
    arquivos = listar_arquivos_da_pasta(servico, id_pasta)

    baixados = []
    for arq in arquivos:
        if arq["mimeType"] == "application/vnd.google-apps.folder":
            continue

        caminho_local = os.path.join(destino_local, arq["name"])

        if arq["mimeType"] == "application/vnd.google-apps.spreadsheet":
            if not caminho_local.lower().endswith(".xlsx"):
                caminho_local += ".xlsx"
            request = servico.files().export_media(
                fileId=arq["id"],
                mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            request = servico.files().get_media(fileId=arq["id"])

        fh = io.FileIO(caminho_local, "wb")
        downloader = MediaIoBaseDownload(fh, request)
        concluido = False
        while not concluido:
            _, concluido = downloader.next_chunk()
        fh.close()
        baixados.append(os.path.basename(caminho_local))

    return baixados


def subir_arquivos_novos(servico, id_pasta, pasta_local, arquivos_originais, apenas_novos=False):
    """
    Sobe para o Drive os arquivos criados/atualizados pelos robôs na pasta temporária.
    Se `apenas_novos=True`, ignora os arquivos de entrada que não são relatórios finais.
    """
    arquivos_existentes_no_drive = {
        a["name"]: a["id"] for a in listar_arquivos_da_pasta(servico, id_pasta)
    }
    originais_set = set(arquivos_originais or [])

    enviados = []
    for nome_arquivo in os.listdir(pasta_local):
        caminho_completo = os.path.join(pasta_local, nome_arquivo)
        if not os.path.isfile(caminho_completo):
            continue
        if nome_arquivo.startswith("~$") or nome_arquivo.endswith("_temp_conv.xlsx"):
            continue

        if apenas_novos and (nome_arquivo in originais_set) and not (
            nome_arquivo.startswith("AUT")
            or nome_arquivo.startswith("CONTROLE")
            or nome_arquivo.startswith("CONCILIACAO")
            or nome_arquivo.startswith("COMPROVANTES")
            or nome_arquivo.startswith("LOG_ERROS")
        ):
            continue

        media = MediaFileUpload(caminho_completo, resumable=True)

        if nome_arquivo in arquivos_existentes_no_drive:
            servico.files().update(
                fileId=arquivos_existentes_no_drive[nome_arquivo],
                media_body=media,
            ).execute()
        else:
            metadata = {"name": nome_arquivo, "parents": [id_pasta]}
            servico.files().create(body=metadata, media_body=media).execute()

        enviados.append(nome_arquivo)

    return enviados
