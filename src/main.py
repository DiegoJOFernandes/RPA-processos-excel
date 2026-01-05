from pathlib import Path

from src.config import settings
from src.io_excel import read_input_excel
from src.transform import validate_and_clean, group_invoices, invoice_header_from_group
from src.fill_template import fill_invoice_template
from src.print_invoice import export_invoice_pdf, print_invoice
from src.preflight import preflight_checks  # ✅ novo import


def main():
    # ===============================
    # 1) Leitura e validação inicial
    # ===============================
    df = read_input_excel()
    df = validate_and_clean(df)

    # ===============================
    # 2) Preflight checks (ANTES do RPA)
    # ===============================
    report = preflight_checks(df)

    print(
        f"[PRECHECK] Linhas: {report.rows}\n"
        f"[PRECHECK] Faturas totais: {report.invoices_total} "
        f"(PF={report.invoices_pf}, PJ={report.invoices_pj})\n"
        f"[PRECHECK] Input: {report.input_path.resolve()}\n"
        f"[PRECHECK] Template PF: {report.template_pf.resolve()}\n"
        f"[PRECHECK] Template PJ: {report.template_pj.resolve()}\n"
        f"[PRECHECK] Output: {report.output_root.resolve()}\n"
    )

    output_root = Path(settings.output_dir)

    # ===============================
    # 3) Loop por cliente (grupo)
    # ===============================
    for doc, group in group_invoices(df):

        # Identifica PF ou PJ
        client_type = str(
            group[settings.client_type_column.lower()].iloc[0]
        ).strip().upper()

        # Escolhe o template correto
        if client_type == "PF":
            template_file = Path(settings.template_pf)
        elif client_type == "PJ":
            template_file = Path(settings.template_pj)
        else:
            raise ValueError(f"Tipo de cliente inválido: {client_type}")

        # Segurança extra (redundante ao preflight, mas ok)
        if not template_file.exists():
            raise FileNotFoundError(
                f"Template não encontrado: {template_file.resolve()}"
            )

        # Monta cabeçalho da fatura
        header = invoice_header_from_group(doc, group)

        # ===============================
        # 4) Monta itens (transações)
        # ===============================
        items = []
        for _, row in group.iterrows():
            valor_compra = str(row.get("valor_compra", "0")).replace(",", ".")
            qtd_parcelas = str(row.get("qtd_parcelas", "1")).replace(",", ".")
            valor_parcela = str(row.get("valor_parcela", "0")).replace(",", ".")

            try:
                valor_compra_f = float(valor_compra)
            except ValueError:
                valor_compra_f = 0.0

            try:
                qtd_parcelas_i = int(float(qtd_parcelas))
            except ValueError:
                qtd_parcelas_i = 1

            try:
                valor_parcela_f = float(valor_parcela)
            except ValueError:
                valor_parcela_f = 0.0

            estabelecimento = str(row.get("estabelecimento", "")).strip()

            items.append({
                "descricao": (
                    f"{estabelecimento} | "
                    f"Compra: R$ {valor_compra_f:.2f} | "
                    f"{qtd_parcelas_i}x"
                ),
                "quantidade": 1,
                "valor_unitario": valor_parcela_f,
                "valor_total": valor_parcela_f,
            })

        # ===============================
        # 5) Saída da fatura
        # ===============================
        invoice_folder = output_root / client_type / f"FATURA_{doc}"
        invoice_folder.mkdir(parents=True, exist_ok=True)

        output_file = invoice_folder / f"fatura_{doc}.xlsx"
        pdf_file = invoice_folder / f"fatura_{doc}.pdf"

        # Preenche XLSX
        fill_invoice_template(
            header=header,
            items=items,
            template_file=template_file,
            output_path=output_file,
        )

        # Exporta PDF
        try:
            export_invoice_pdf(output_file, pdf_file)
            pdf_status = "PDF_OK"
        except Exception as e:
            pdf_status = f"PDF_FAIL: {e}"

        # Impressão (opcional)
        try:
            print_invoice(output_file)
            print_status = "PRINT_OK"
        except Exception as e:
            print_status = f"PRINT_FAIL: {e}"

        # Status por fatura
        status_text = f"{pdf_status}\n{print_status}\n"
        (invoice_folder / "status.txt").write_text(status_text, encoding="utf-8")

    print("Processamento concluído.")


if __name__ == "__main__":
    main()
