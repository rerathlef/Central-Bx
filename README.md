# Central Bx Finance — Automação Multi-Projetos (Cativa & Revpack)

Central unificada no padrão **Bx finance** para rodar os robôs de **Conciliação Bancária** e **Controle de Sacados** de múltiplos grupos empresariais em um único aplicativo privado no Streamlit Cloud.

## Estrutura de Pastas

- `streamlit_app.py` — Painel principal com seletor de projeto (`REVPACK` | `CATIVA`), mantendo caminhos pré-preenchidos independentes e execuções isoladas em pastas temporárias.
- `drive_utils.py` — Conector OAuth 2.0 com os IDs das pastas raiz da Cativa e da Revpack.
- `requirements.txt` — Todas as dependências para PDFs, Excel (.xlsx / .xls / HTML) e Google Drive API.
- `robos/`
  - `CATIVA/`
    - `pdf_pt1.py`
    - `sicoob.py`
    - `outros.py`
  - `REVPACK/`
    - `REV - C.B. - AUT.py`
    - `REVPACK - C.S..py`

## Dica de Link Direto por Projeto

Você pode abrir o site já selecionado no projeto desejado adicionando `?projeto=REVPACK` ou `?projeto=CATIVA` no final do endereço do app.
