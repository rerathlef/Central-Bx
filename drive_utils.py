import io
import os
import unicodedata
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from google.oauth2.credentials import Credentials

SCOPES = ["https://www.googleapis.com/auth/drive"]

ID_PASTA_RAIZ_CATIVA = "1F6NiwgME7xPe_xUQQwzQHwwaBenRXO6N"
ID_PASTA_REVPACK_CONCILIACAO = "1a5CLnMLhinJAAqdD9Tdr3fvFBtcf_hj0"
ID_PASTA_REVPACK_SACADOS = "1a5CLnMLhinJAAqdD9Tdr3fvFBtcf_hj0"
ID_PASTA_RAIZ = ID_PASTA_RAIZ_CATIVA


def _normalizar_nome(texto):
    if not texto:
        return ""
    sem_acento = "".join(
        c for c in unicodedata.normalize("NFKD", str(texto))
        if not unicodedata.combining(c)
    )
    return " ".join(sem_acento.upper().strip().split())


def conectar_drive(credenciais_dict):
    info = dict(credenciais_dict)
    if "expiry" in info:
        del info["expiry"]
    creds = Credentials.from_authorized_user_info(info, SCOPES)
    servico = build("drive", "v3", credentials=creds)
    return servico


def obter_email_autenticado(servico):
    try:
        info = servico.about().get(fields="user(emailAddress, displayName)").execute()
        user = info.get("user", {})
        email = user.get("emailAddress", "desconhecido")
        nome = user.get("displayName", "")
        return f"{nome} <{email}>".strip() if nome else email
    except Exception:
        return "Conta nao identificada"


def _resolver_id_real(item):
    if item.get("mimeType") == "application/vnd.google-apps.shortcut":
        detalhes = item.get("shortcutDetails", {})
        target_id = detalhes.get("targetId")
        if target_id:
            return target_id
    return item["id"]


def _listar_subpastas_pai(servico, id_pasta_pai):
    pastas = []
    page_token = None
    query = (
        f"'{id_pasta_pai}' in parents "
        f"and (mimeType = 'application/vnd.google-apps.folder' "
        f"     or mimeType = 'application/vnd.google-apps.shortcut') "
        f"and trashed = false"
    )
    while True:
        res = servico.files().list(
            q=query,
            fields="nextPageToken, files(id, name, mimeType, shortcutDetails)",
            pageToken=page_token,
            pageSize=100,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        pastas.extend(res.get("files", []))
        page_token = res.get("nextPageToken")
        if not page_token:
            break
    return pastas


def _achar_subpasta(servico, nome_pasta, id_pasta_pai):
    nome_pai = id_pasta_pai
    try:
        meta_pai = servico.files().get(
            fileId=id_pasta_pai,
            fields="id, name, mimeType, shortcutDetails",
            supportsAllDrives=True,
        ).execute()
        nome_pai = f"'{meta_pai.get('name', id_pasta_pai)}' (id={id_pasta_pai})"
        if meta_pai.get("mimeType") == "application/vnd.google-apps.shortcut":
            id_pasta_pai = meta_pai.get("shortcutDetails", {}).get("targetId", id_pasta_pai)
    except Exception as e:
        email_atual = obter_email_autenticado(servico)
        raise FileNotFoundError(
            f"A conta do token ({email_atual}) NAO tem permissao na pasta pai (id={id_pasta_pai}). "
            f"Compartilhe essa pasta no Google Drive com {email_atual}. Detalhe: {e}"
        )

    subpastas = _listar_subpastas_pai(servico, id_pasta_pai)
    alvo_norm = _normalizar_nome(nome_pasta)

    for item in subpastas:
        if item["name"].strip() == nome_pasta.strip():
            return _resolver_id_real(item)

    for item in subpastas:
        if _normalizar_nome(item["name"]) == alvo_norm:
            return _resolver_id_real(item)

    nomes_visiveis = [p["name"] for p in subpastas]
    email_atual = obter_email_autenticado(servico)
    if not nomes_visiveis:
        raise FileNotFoundError(
            f"Pasta '{nome_pasta}' nao encontrada dentro de {nome_pai}. "
            f"Nenhuma subpasta visivel para {email_atual} (verifique o compartilhamento no Drive)."
        )
    raise FileNotFoundError(
        f"Pasta '{nome_pasta}' nao encontrada dentro de {nome_pai}. "
        f"Subpastas visiveis neste local: {nomes_visiveis}"
    )


def navegar_caminho(servico, caminho_relativo, id_raiz=ID_PASTA_RAIZ):
    caminho_normalizado = caminho_relativo.replace("\\", "/")
    partes = [p.strip() for p in caminho_normalizado.split("/") if p.strip()]
    id_atual = id_raiz

    nome_raiz_norm = ""
    try:
        meta_raiz = servico.files().get(
            fileId=id_raiz,
            fields="id, name, mimeType, shortcutDetails",
            supportsAllDrives=True,
        ).execute()
        nome_raiz_norm = _normalizar_nome(meta_raiz.get("name", ""))
        if meta_raiz.get("mimeType") == "application/vnd.google-apps.shortcut":
            id_atual = meta_raiz.get("shortcutDetails", {}).get("targetId", id_raiz)
    except Exception:
        email_atual = obter_email_autenticado(servico)
        raise FileNotFoundError(
            f"A conta do token ({email_atual}) NAO tem acesso a pasta raiz (id={id_raiz}). "
            f"Compartilhe a pasta no Google Drive com {email_atual}."
        )

    for idx, nome_parte in enumerate(partes):
        parte_norm = _normalizar_nome(nome_parte)
        if idx == 0 and (parte_norm == nome_raiz_norm or parte_norm == "12 CATIVA"):
            continue
        id_atual = _achar_subpasta(servico, nome_parte, id_atual)

    return id_atual


def listar_arquivos_da_pasta(servico, id_pasta):
    arquivos = []
    page_token = None
    while True:
        query = f"'{id_pasta}' in parents and trashed = false"
        resposta = servico.files().list(
            q=query,
            fields="nextPageToken, files(id, name, mimeType, shortcutDetails)",
            pageToken=page_token,
            pageSize=100,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        arquivos.extend(resposta.get("files", []))
        page_token = resposta.get("nextPageToken")
        if not page_token:
            break
    return arquivos


def baixar_pasta_inteira(servico, id_pasta, destino_local):
    os.makedirs(destino_local, exist_ok=True)
    arquivos = listar_arquivos_da_pasta(servico, id_pasta)
    baixados = []

    for arq in arquivos:
        mime = arq.get("mimeType", "")
        file_id = arq["id"]

        if mime == "application/vnd.google-apps.shortcut":
            detalhes = arq.get("shortcutDetails", {})
            target_mime = detalhes.get("targetMimeType", "")
            target_id = detalhes.get("targetId")
            if target_id:
                file_id = target_id
                mime = target_mime

        if mime == "application/vnd.google-apps.folder":
            continue

        caminho_local = os.path.join(destino_local, arq["name"])

        if mime == "application/vnd.google-apps.spreadsheet":
            if not caminho_local.lower().endswith(".xlsx"):
                caminho_local += ".xlsx"
            request = servico.files().export_media(
                fileId=file_id,
                mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            request = servico.files().get_media(fileId=file_id, supportsAllDrives=True)

        fh = io.FileIO(caminho_local, "wb")
        downloader = MediaIoBaseDownload(fh, request)
        concluido = False
        while not concluido:
            _, concluido = downloader.next_chunk()
        fh.close()
        baixados.append(os.path.basename(caminho_local))

    return baixados


def baixar_arquivo_controle_oficial(servico, nome_arquivo, destino_local, id_pasta_extratos, id_pasta_financeiro=""):
    """
    Localiza o arquivo oficial (ex: CONTROLE DE SACADOS.xlsx) na pasta 03 Financeiro
    (ou busca pelo nome no Drive) e baixa para destino_local. Retorna (id_arquivo_drive, caminho_local).
    """
    pastas_para_procurar = []
    if id_pasta_financeiro:
        pastas_para_procurar.append(id_pasta_financeiro)

    # Descobre automaticamente a pasta pai de '02. EXTRATOS' (que é a '03 Financeiro')
    if id_pasta_extratos:
        try:
            meta = servico.files().get(
                fileId=id_pasta_extratos,
                fields="parents",
                supportsAllDrives=True,
            ).execute()
            parents = meta.get("parents", [])
            for p in parents:
                if p not in pastas_para_procurar:
                    pastas_para_procurar.append(p)
        except Exception:
            pass
        if id_pasta_extratos not in pastas_para_procurar:
            pastas_para_procurar.append(id_pasta_extratos)

    alvo_norm = _normalizar_nome(nome_arquivo)
    arquivo_encontrado = None

    # 1. Procura nas pastas candidatas (03 Financeiro / 02. EXTRATOS)
    for id_p in pastas_para_procurar:
        try:
            itens = listar_arquivos_da_pasta(servico, id_p)
            for arq in itens:
                if _normalizar_nome(arq["name"]) == alvo_norm:
                    arquivo_encontrado = arq
                    break
            if arquivo_encontrado:
                break
        except Exception:
            continue

    # 2. Fallback: se não achou na pasta pai, procura pelo nome exato no Drive
    if not arquivo_encontrado:
        nome_escaped = nome_arquivo.replace("'", "\\'")
        query = f"name = '{nome_escaped}' and trashed = false"
        res = servico.files().list(
            q=query,
            fields="files(id, name, mimeType, shortcutDetails)",
            pageSize=10,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        arquivos = res.get("files", [])
        if arquivos:
            arquivo_encontrado = arquivos[0]

    if not arquivo_encontrado:
        raise FileNotFoundError(
            f"Arquivo oficial '{nome_arquivo}' nao foi encontrado na pasta '03 Financeiro'. "
            f"Verifique se o arquivo '{nome_arquivo}' esta compartilhado no Google Drive."
        )

    file_id = _resolver_id_real(arquivo_encontrado)
    caminho_local = os.path.join(destino_local, nome_arquivo)

    request = servico.files().get_media(fileId=file_id, supportsAllDrives=True)
    fh = io.FileIO(caminho_local, "wb")
    downloader = MediaIoBaseDownload(fh, request)
    concluido = False
    while not concluido:
        _, concluido = downloader.next_chunk()
    fh.close()

    return file_id, caminho_local


def atualizar_arquivo_por_id(servico, file_id, caminho_local):
    """Sobrescreve um arquivo existente no Google Drive mantendo o mesmo ID."""
    media = MediaFileUpload(caminho_local, resumable=True)
    servico.files().update(
        fileId=file_id,
        media_body=media,
        supportsAllDrives=True,
    ).execute()


def subir_arquivos_novos(servico, id_pasta, pasta_local, arquivos_originais, apenas_novos=False):
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
                supportsAllDrives=True,
            ).execute()
        else:
            metadata = {"name": nome_arquivo, "parents": [id_pasta]}
            servico.files().create(
                body=metadata,
                media_body=media,
                supportsAllDrives=True,
            ).execute()

        enviados.append(nome_arquivo)

    return enviados
