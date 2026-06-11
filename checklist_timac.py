"""
checklist_timac.py
Preenche a planilha Check_list_de_Veículos.xls/.xlsx com os dados da ordem TIMAC.
Uso:
    from checklist_timac import preencher_checklist_timac
    preencher_checklist_timac(timac_dados, caminho_planilha)
"""

import os
import shutil
from pathlib import Path


TIPOS_STR = {
    "TRUCK":           "TIPO: (  X  ) TRUCK    (   ) BI-TRUCK   (   ) CARRETA SIMPLES   (    ) TRUCADA   (    ) VANDERLÉIA   (    ) BI-TREM   (    ) RODO-TREM    (    ) CAÇAMBA ",
    "BI-TRUCK":        "TIPO: (   ) TRUCK    (  X  ) BI-TRUCK   (   ) CARRETA SIMPLES   (    ) TRUCADA   (    ) VANDERLÉIA   (    ) BI-TREM   (    ) RODO-TREM    (    ) CAÇAMBA ",
    "CARRETA SIMPLES": "TIPO: (   ) TRUCK    (   ) BI-TRUCK   (  X  ) CARRETA SIMPLES   (    ) TRUCADA   (    ) VANDERLÉIA   (    ) BI-TREM   (    ) RODO-TREM    (    ) CAÇAMBA ",
    "TRUCADA":         "TIPO: (   ) TRUCK    (   ) BI-TRUCK   (   ) CARRETA SIMPLES   (   X ) TRUCADA   (    ) VANDERLÉIA   (    ) BI-TREM   (    ) RODO-TREM    (    ) CAÇAMBA ",
    "VANDERLÉIA":      "TIPO: (   ) TRUCK    (   ) BI-TRUCK   (   ) CARRETA SIMPLES   (    ) TRUCADA   (   X ) VANDERLÉIA   (    ) BI-TREM   (    ) RODO-TREM    (    ) CAÇAMBA ",
    "BI-TREM":         "TIPO: (   ) TRUCK    (   ) BI-TRUCK   (   ) CARRETA SIMPLES   (    ) TRUCADA   (    ) VANDERLÉIA   (   X ) BI-TREM   (    ) RODO-TREM    (    ) CAÇAMBA ",
    "RODO-TREM":       "TIPO: (   ) TRUCK    (   ) BI-TRUCK   (   ) CARRETA SIMPLES   (    ) TRUCADA   (    ) VANDERLÉIA   (    ) BI-TREM   (   X ) RODO-TREM    (    ) CAÇAMBA ",
    "CAÇAMBA":         "TIPO: (   ) TRUCK    (   ) BI-TRUCK   (   ) CARRETA SIMPLES   (    ) TRUCADA   (    ) VANDERLÉIA   (    ) BI-TREM   (    ) RODO-TREM    (   X ) CAÇAMBA ",
}

ESTADOS_STR = {
    "LÍQUIDO": "PRODUTO:  LÍQUIDO (  X  )    SÓLIDO (   )     PÓ (   )",
    "SÓLIDO":  "PRODUTO:  LÍQUIDO (   )    SÓLIDO (  X  )     PÓ (   )",
    "PÓ":      "PRODUTO:  LÍQUIDO (   )    SÓLIDO (   )     PÓ (  X  )",
    "":        "PRODUTO:  LÍQUIDO (   )    SÓLIDO (   )     PÓ (   )",
}


def _modelo_path():
    """Localiza a planilha modelo no mesmo diretório do executável/script."""
    import sys
    if getattr(sys, "frozen", False):
        # Executável compilado — procura na pasta do .exe
        base = Path(sys.executable).parent
    else:
        base = Path(__file__).parent
    for nome in ["Check_list_de_Veículos.xls", "Check_list_de_Veiculos.xls",
                 "Check_list_de_Veículos.xlsx", "checklist_veiculos.xlsx"]:
        p = base / nome
        if p.exists():
            return p
    return None


def preencher_checklist_timac(dados: dict, pasta_destino: str) -> str:
    import sys, shutil, os, time as _time
    from pathlib import Path
    import xlwings as xw

    # Log
    try:
        _base_log = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
        _logf = str(_base_log / "checklist_log.txt")
        with open(_logf, "a", encoding="utf-8") as _lf:
            import datetime as _dtt
            _lf.write(f"\n[{_dtt.datetime.now()}] INICIO\n")
            _lf.write(f"  pasta_destino={pasta_destino}\n")
    except Exception:
        pass

    modelo = _modelo_path()

    try:
        with open(_logf, "a", encoding="utf-8") as _lf:
            _lf.write(f"  modelo={modelo}\n")
    except Exception:
        pass

    if not modelo:
        raise FileNotFoundError(
            "Planilha modelo 'Check_list_de_Veiculos.xlsx' nao encontrada na pasta do sistema."
        )

    # Nome no formato CL_motorista_placa
    motorista  = dados.get("motorista", "MOTORISTA").upper().replace(" ", "_")
    placa_nome = dados.get("placa_cavalo", "").upper().replace("-", "").replace(" ", "")
    nome_xlsx  = f"CL_{motorista}_{placa_nome}.xlsx"
    nome_pdf   = f"CL_{motorista}_{placa_nome}.pdf"
    caminho_saida = str(Path(pasta_destino) / nome_xlsx)
    caminho_pdf   = str(Path(pasta_destino) / nome_pdf)

    shutil.copy(str(modelo), caminho_saida)

    # Dados
    data          = dados.get("data", "")
    motorista_nom = dados.get("motorista", "")
    rg            = dados.get("rg", "")
    cpf           = dados.get("cpf", "")
    nº_cnh        = dados.get("nº_cnh", "")
    categoria_cnh = dados.get("categoria_cnh", "")
    validade_cnh  = dados.get("validade_cnh", "")
    produto       = dados.get("produto", "")
    estado        = dados.get("estado_fisico", "").upper()
    tipo_veiculo  = dados.get("tipo_veiculo", "").upper()
    placa_cavalo  = dados.get("placa_cavalo", "")
    carreta1      = dados.get("carreta1", "")
    carreta2      = dados.get("carreta2", "")
    carreta3      = dados.get("carreta3", "")

    carretas = " / ".join(c for c in [carreta1, carreta2, carreta3] if c)
    str_a12 = (
        f"CIV CARRETA:          CIPP CARRETA:          CIV CAVALO:          "
        f"PLACA DO CAVALO: {placa_cavalo}       "
        f"PLACA DA CARRETA: {carretas}"
    )

    app_xl = None
    wb_xl  = None
    xl_fallback = None

    def _fechar_tudo():
        nonlocal wb_xl, app_xl, xl_fallback
        for obj, metodo in [(wb_xl, "close"), (app_xl, "quit")]:
            if obj is not None:
                try: getattr(obj, metodo)()
                except Exception: pass
        wb_xl = app_xl = None
        if xl_fallback is not None:
            try: xl_fallback.Quit()
            except Exception: pass
            xl_fallback = None

    try:
        app_xl = xw.App(visible=False)
        app_xl.api.Visible = False
        app_xl.api.ScreenUpdating = False
        app_xl.api.DisplayAlerts = False
        app_xl.display_alerts = False

        wb_xl = xw.Book(os.path.abspath(caminho_saida))
        ws = wb_xl.sheets[0]

        # Preenche células
        ws["C7"].value  = f"DATA: {data}"
        ws["A8"].value  = f"NOME DO MOTORISTA: {motorista_nom}"
        ws["C8"].value  = f"RG: {rg}"
        ws["E8"].value  = f"CPF: {cpf}"
        ws["A9"].value  = f"NUMERO CNH: {nº_cnh}"
        ws["B9"].value  = f"CATEGORIA DA CNH: {categoria_cnh}"
        ws["C9"].value  = f"VALIDADE: {validade_cnh}"
        ws["A10"].value = ESTADOS_STR.get(estado, ESTADOS_STR[""])
        if produto:
            ws["C10"].value = f"DESCRICAO DO PRODUTO: {produto}"
        ws["A11"].value = TIPOS_STR.get(tipo_veiculo, TIPOS_STR.get("TRUCK", ""))
        ws["A12"].value = str_a12

        wb_xl.save()

        # Exporta PDF
        abs_pdf = os.path.abspath(caminho_pdf)
        try:
            wb_xl.api.ExportAsFixedFormat(0, abs_pdf)
        except Exception:
            pass

        # Fallback win32com
        if not os.path.exists(abs_pdf) or os.path.getsize(abs_pdf) == 0:
            try:
                wb_xl.close(); wb_xl = None
                app_xl.quit();  app_xl = None
            except Exception:
                pass
            _time.sleep(1)
            try:
                import win32com.client as _win32
                xl_fallback = _win32.Dispatch("Excel.Application")
                xl_fallback.Visible = False
                xl_fallback.DisplayAlerts = False
                wb2 = xl_fallback.Workbooks.Open(os.path.abspath(caminho_saida))
                try:
                    wb2.ExportAsFixedFormat(0, abs_pdf)
                finally:
                    wb2.Close(False)
            except Exception as e_fb:
                raise RuntimeError(f"Nao foi possivel gerar o PDF do checklist: {e_fb}")
    finally:
        _fechar_tudo()

    try:
        os.remove(caminho_saida)
    except Exception:
        pass

    if not os.path.exists(caminho_pdf) or os.path.getsize(caminho_pdf) == 0:
        raise RuntimeError("PDF do checklist nao foi gerado.")

    return caminho_pdf