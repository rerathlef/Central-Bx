import os
import sys
import pandas as pd
import openpyxl
from copy import copy

# ==============================================================================
# CONFIGURAÇÃO - SUPORTE AUTOMÁTICO PARA STREAMLIT (NUVEM) E SPYDER (LOCAL)
# ==============================================================================
NOME_ABA_BASE = "BASE"

if len(sys.argv) >= 4 and not sys.argv[1].startswith("-"):
    # Modo Nuvem / Streamlit: recebe (CAMINHO_P1, CAMINHO_P2, DATA_HISTORICO)
    CAMINHO_P1 = sys.argv[1]
    CAMINHO_P2 = sys.argv[2]
    DATA_HISTORICO_ARG = sys.argv[3].strip()
else:
    # Modo Local (Spyder)
    CAMINHO_P1 = r"G:\.shortcut-targets-by-id\1HH8YPWOEhDI7x9lk-LNQXlVf9sUusDFg\08 Revpack\03 Financeiro\02. EXTRATOS\POSICOES FUNDOS\2026\10 - OUT\Fundos 13.10\Trabalhados\teste robo 1\TESTADO\CONTROLE DE SACADOS AUTOMATIZADO - REVPACK.xlsx"
    CAMINHO_P2 = r"G:\.shortcut-targets-by-id\1HH8YPWOEhDI7x9lk-LNQXlVf9sUusDFg\08 Revpack\03 Financeiro\CONTROLE DE SACADOS.xlsx"
    DATA_HISTORICO_ARG = None


def perguntar_data_local():
    """Usado apenas se rodar localmente no Spyder sem argumentos."""
    try:
        import tkinter as tk
        from tkinter import simpledialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        data_input = simpledialog.askstring(
            "Atualização de Histórico",
            "Qual a data da ÚLTIMA atualização? (Ex: 25-11)\nIsso será o nome da aba de backup.",
        )
        root.destroy()
        return data_input
    except Exception:
        return input("Qual a data da ÚLTIMA atualização? (Ex: 25-11): ").strip()


def executar_atualizacao():
    print("--- 🚀 INICIANDO ATUALIZAÇÃO DO CONTROLE DE SACADOS PT 2 (MODO HEADLESS) ---")

    if not os.path.exists(CAMINHO_P1):
        print(f"❌ ERRO: Arquivo P1 não encontrado: {CAMINHO_P1}")
        sys.exit(1)

    if not os.path.exists(CAMINHO_P2):
        print(f"❌ ERRO: Arquivo P2 (Controle Oficial) não encontrado: {CAMINHO_P2}")
        sys.exit(1)

    data_historico = DATA_HISTORICO_ARG if DATA_HISTORICO_ARG else perguntar_data_local()
    if not data_historico:
        print("⚠️ Operação cancelada: nenhuma data de histórico informada.")
        sys.exit(1)

    nome_aba_historico = f"BASE {data_historico}".strip()

    try:
        print("-> Carregando P2 (CONTROLE DE SACADOS.xlsx) em segundo plano...")
        wb_p2 = openpyxl.load_workbook(CAMINHO_P2)

        if NOME_ABA_BASE not in wb_p2.sheetnames:
            print(f"❌ ERRO: Aba '{NOME_ABA_BASE}' não encontrada no arquivo oficial. Abas existentes: {wb_p2.sheetnames}")
            sys.exit(1)

        sheet_base = wb_p2[NOME_ABA_BASE]

        # 3. Criar Histórico (Copiar aba BASE e posicionar antes da aba BASE)
        print(f"-> Criando cópia de segurança '{nome_aba_historico}'...")
        if nome_aba_historico in wb_p2.sheetnames:
            print(f"   ⚠️ Aviso: A aba '{nome_aba_historico}' já existia e será substituída pela nova cópia.")
            del wb_p2[nome_aba_historico]

        sheet_nova = wb_p2.copy_worksheet(sheet_base)
        sheet_nova.title = nome_aba_historico

        # Move a aba de backup para ficar imediatamente ANTES da aba BASE (igual ao before=sheet_base do xlwings)
        abas = wb_p2._sheets
        abas.remove(sheet_nova)
        idx_base = abas.index(sheet_base)
        abas.insert(idx_base, sheet_nova)

        # 4. Limpar dados antigos da aba BASE (Linha 6 para baixo, mantendo fórmulas nas colunas D, H, J, etc.)
        print("-> Limpando dados antigos da aba BASE (Linha 6 para baixo, preservando fórmulas)...")
        last_row = sheet_base.max_row
        for r in range(sheet_base.max_row, 5, -1):
            if sheet_base.cell(row=r, column=2).value not in (None, ""):
                last_row = r
                break

        colunas_limpar = [2, 3, 5, 6, 7, 9, 11, 12]  # B, C, E, F, G, I, K, L
        if last_row >= 6:
            for r in range(6, last_row + 1):
                for col_idx in colunas_limpar:
                    sheet_base.cell(row=r, column=col_idx).value = None

        # 5. Ler os dados novos da P1
        print("-> Lendo dados novos gerados pelo Robô P1...")
        df_p1 = pd.read_excel(CAMINHO_P1)
        total_linhas = len(df_p1)
        print(f"   -> {total_linhas} linhas carregadas do P1.")

        # 6. Colar na P2 conforme mapeamento exato (a partir da linha 6)
        print("-> Inserindo dados novos na aba BASE da P2...")
        # Mapeamento de colunas: (índice 0-based no df_p1) -> (coluna 1-based no Excel P2)
        # P1 A,B (0,1) -> P2 B,C (2,3)
        # P1 C,D,E (2,3,4) -> P2 E,F,G (5,6,7)
        # P1 F (5) -> P2 I (9)
        # P1 G,H (6,7) -> P2 K,L (11,12)
        mapa_colunas = {
            0: 2,   # A -> B
            1: 3,   # B -> C
            2: 5,   # C -> E
            3: 6,   # D -> F
            4: 7,   # E -> G
            5: 9,   # F -> I
            6: 11,  # G -> K
            7: 12,  # H -> L
        }

        valores_matriz = df_p1.values.tolist()
        for offset, linha_vals in enumerate(valores_matriz):
            excel_row = 6 + offset
            for idx_p1, col_p2 in mapa_colunas.items():
                if idx_p1 < len(linha_vals):
                    val = linha_vals[idx_p1]
                    if pd.isna(val):
                        val = None
                    sheet_base.cell(row=excel_row, column=col_p2).value = val

        print("-> Salvando CONTROLE DE SACADOS.xlsx...")
        wb_p2.save(CAMINHO_P2)
        wb_p2.close()
        print(f"✅ SUCESSO! Aba de histórico '{nome_aba_historico}' criada e aba 'BASE' atualizada com {total_linhas} títulos!")

    except Exception as e:
        print(f"❌ ERRO CRÍTICO NA PARTE 2: {e}")
        sys.exit(1)


if __name__ == "__main__":
    executar_atualizacao()