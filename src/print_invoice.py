from __future__ import annotations

from pathlib import Path
import platform


def _ensure_windows() -> None:
    if platform.system() != "Windows":
        raise RuntimeError(
            "Exportação/Impressão automática requer Windows com Microsoft Excel instalado."
        )


def export_invoice_pdf_windows(xlsx_path: Path, pdf_path: Path) -> None:
    """
    Abre um arquivo .xlsx no Excel e exporta a planilha 1 para PDF em tamanho A4.

    Requisitos:
      - Windows
      - Microsoft Excel instalado
      - pywin32 instalado

    Args:
      xlsx_path: caminho do arquivo .xlsx já gerado (fatura)
      pdf_path: caminho de saída do PDF (ex: output/FATURA_xxx/fatura_xxx.pdf)
    """
    _ensure_windows()

    import win32com.client  # type: ignore

    xlsx_path = xlsx_path.resolve()
    pdf_path = pdf_path.resolve()
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    # Constantes do Excel (valores numéricos para evitar dependência de enums)
    XL_TYPE_PDF = 0
    XL_QUALITY_STANDARD = 0
    # Paper size A4 no Excel: xlPaperA4 = 9
    XL_PAPER_A4 = 9

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(str(xlsx_path))
        try:
            ws = wb.Worksheets(1)

            # Configuração de página para A4 + ajustes de impressão
            ps = ws.PageSetup
            ps.PaperSize = XL_PAPER_A4
            ps.Orientation = 1  # 1 = Portrait, 2 = Landscape (ajuste se preferir)
            ps.Zoom = False
            ps.FitToPagesWide = 1   # encaixar em 1 página na largura
            ps.FitToPagesTall = False  # não força altura (deixa quebrar se necessário)

            # Exporta para PDF (planilha ativa / primeira planilha)
            ws.ExportAsFixedFormat(
                Type=XL_TYPE_PDF,
                Filename=str(pdf_path),
                Quality=XL_QUALITY_STANDARD,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


def print_excel_windows(xlsx_path: Path) -> None:
    """
    Imprime a primeira planilha de um arquivo .xlsx via Excel (Windows).
    """
    _ensure_windows()

    import win32com.client  # type: ignore

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(str(xlsx_path.resolve()))
        try:
            wb.Worksheets(1).PrintOut()
        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


def export_invoice_pdf(xlsx_path: Path, pdf_path: Path) -> None:
    """
    Wrapper multiplataforma (por enquanto apenas Windows).
    """
    export_invoice_pdf_windows(xlsx_path, pdf_path)


def print_invoice(xlsx_path: Path) -> None:
    """
    Wrapper multiplataforma (por enquanto apenas Windows).
    """
    print_excel_windows(xlsx_path)
