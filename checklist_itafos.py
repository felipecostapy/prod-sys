"""
checklist_itafos.py
Preenche a ORDEM INDIVIDUAL DE CARREGAMENTO - ITAFOS com os dados da ordem.

Mapeamento de células (aba "O.C"):
  B11        → CLIENTE
  K11        → PEDIDO
  C14        → MOTORISTA
  I14        → CPF
  B15        → PESO BRUTO  (calculado: não salvo no banco)
  E15        → CAVALO
  H15        → CARRETA 1 + CARRETA 2  (concatenadas)
  K15        → CARRETA 3
  C16        → DATA AGENDAMENTO
  A19        → TONELADAS A CARREGAR
  E19        → PRODUTO
  L16        → TIPO DE EMBALAGEM  (GRANEL / BIG BAG / TANQUE)
"""

import sys
import re
import datetime
from pathlib import Path

# ── Compatibilidade frozen (PyInstaller) ─────────────────────────
def _modelo_path() -> Path:
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).parent
    else:
        base = Path(__file__).parent
    return base / "ORDEM_INDIVIDUAL_DE_CARREGAMENTO_-_ITAFOS.xlsx"


def _normalizar_placa(v: str) -> str:
    s = re.sub(r"[^A-Za-z0-9]", "", str(v or "")).upper()
    if len(s) == 7:
        return s[:3] + "-" + s[3:]
    return s


def preencher_itafos(dados: dict, pasta_destino: str | Path | None = None) -> Path:
    """
    Preenche o modelo ITAFOS com os dados da ordem e salva o PDF via xlwings.

    Parâmetros
    ----------
    dados : dict  com as chaves:
        cliente, pedido, motorista, cpf,
        peso_bruto, peso (toneladas),
        placa (cavalo), carreta1, carreta2, carreta3,
        data (str YYYY-MM-DD ou date),
        produto, embalagem (GRANEL / BIG BAG / TANQUE)
    pasta_destino : Path ou None  (padrão: mesma pasta do modelo)

    Retorna
    -------
    Path  do PDF gerado
    """
    try:
        import xlwings as xw
    except ImportError:
        raise RuntimeError("xlwings não instalado. Execute: pip install xlwings")

    modelo = _modelo_path()
    if not modelo.exists():
        raise FileNotFoundError(f"Modelo não encontrado: {modelo}")

    # Formatar data
    data_raw = dados.get("data", "")
    if isinstance(data_raw, (datetime.date, datetime.datetime)):
        data_fmt = data_raw.strftime("%d/%m/%Y")
    elif data_raw:
        # Suporta DD/MM/YYYY (já formatado) ou YYYY-MM-DD (banco)
        if "/" in str(data_raw):
            data_fmt = str(data_raw)  # já está em DD/MM/YYYY
        else:
            try:
                data_fmt = datetime.datetime.strptime(str(data_raw), "%Y-%m-%d").strftime("%d/%m/%Y")
            except ValueError:
                data_fmt = str(data_raw)
    else:
        data_fmt = datetime.date.today().strftime("%d/%m/%Y")

    # Placas
    cavalo   = _normalizar_placa(dados.get("placa",    ""))
    carreta1 = _normalizar_placa(dados.get("carreta1", ""))
    carreta2 = _normalizar_placa(dados.get("carreta2", ""))
    carreta3 = _normalizar_placa(dados.get("carreta3", ""))

    # Carreta 1 + 2 concatenadas
    carr_12_parts = [p for p in [carreta1, carreta2] if p]
    carr_12 = " / ".join(carr_12_parts) if carr_12_parts else ""

    # Toneladas
    try:
        raw = str(dados.get("peso", "") or "0").strip()
        # Se tem vírgula: formato BR (1.234,56) — remove pontos, troca vírgula por ponto
        if "," in raw:
            raw = raw.replace(".", "").replace(",", ".")
        else:
            # Só tem ponto: verifica se é separador de milhar (ex: 50.000, 1.000)
            # Heurística: parte após o ponto tem 3 dígitos → milhar
            partes = raw.split(".")
            if len(partes) == 2 and len(partes[1]) == 3:
                raw = partes[0]  # descarta a parte decimal (era milhar)
        toneladas = float(raw)
        ton_str   = f"{toneladas:g}".replace(".", ",")
    except (ValueError, TypeError):
        ton_str = str(dados.get("peso", ""))

    # Peso bruto (campo calculado, apenas para impressão)
    peso_bruto = str(dados.get("peso_bruto", "")).strip()

    # Produto
    produto = str(dados.get("produto", "") or "SUPERFORTE GRAN GRANEL").upper().strip()

    # Tipo de embalagem — opções: GRANEL / BIG BAG / TANQUE
    embalagem = str(dados.get("embalagem", "GRANEL")).upper().strip()

    # Motorista e cliente em maiúsculas
    cliente   = str(dados.get("cliente",   "") or "").upper().strip()
    pedido    = str(dados.get("pedido",    "") or "").upper().strip()
    motorista = str(dados.get("motorista", "") or "").upper().strip()
    cpf       = str(dados.get("cpf",       "") or "").strip()

    # ── Abrir com xlwings ────────────────────────────────────────
    app = xw.App(visible=False)
    try:
        wb  = app.books.open(str(modelo))
        ws  = wb.sheets["O.C"]

        # Preencher células
        ws["B11"].value = cliente
        ws["K11"].value = pedido
        ws["C14"].value = motorista
        ws["I14"].value = cpf
        ws["B15"].value = peso_bruto
        ws["E15"].value = cavalo
        ws["H15"].value = carr_12
        ws["K15"].value = carreta3
        ws["C16"].value = f"'{data_fmt}"  # apóstrofo força texto puro no Excel
        ws["L16"].value = embalagem
        ws["A19"].value = ton_str
        ws["E19"].value = produto

        # ── Salvar PDF ───────────────────────────────────────────
        if pasta_destino is None:
            pasta_destino = modelo.parent

        pasta_destino = Path(pasta_destino)
        pasta_destino.mkdir(parents=True, exist_ok=True)

        motorista_slug = re.sub(r"[^A-Za-z0-9]", "_", motorista)[:20]
        nome_pdf       = f"OC_ITAFOS_{motorista_slug}_{cavalo}.pdf"
        caminho_pdf    = pasta_destino / nome_pdf

        ws.api.ExportAsFixedFormat(0, str(caminho_pdf))
        wb.close()
        return caminho_pdf

    finally:
        app.quit()


# ── Execução direta (teste) ──────────────────────────────────────
if __name__ == "__main__":
    dados_teste = {
        "cliente":    "FAZENDA MODELO LTDA",
        "pedido":     "39268",
        "motorista":  "EROILDO ALVES DE SOUSA",
        "cpf":        "811.793.081-15",
        "peso_bruto": "34.500",
        "peso":       "33.0",
        "placa":      "FNE8D81",
        "carreta1":   "RBW6E63",
        "carreta2":   "RBW6E93",
        "carreta3":   "RBW6E53",
        "data":       "2026-07-09",
        "produto":    "SUPERFORTE GRAN GRANEL",
        "embalagem":  "GRANEL",
    }
    pdf = preencher_itafos(dados_teste)
    print(f"PDF gerado: {pdf}")