"""
Módulo: docx_pdf_compiler.py (Compilador Direto e Seguro)
Autor: Aristófanes Paulo Gonfulo
Descrição: Atualiza o Sumário com o motor nativo do Word e compila o PDF vetorial.
"""

import os
import win32com.client

def compile_pdf(input_docx, output_pdf):
    abs_docx = os.path.abspath(input_docx)
    abs_pdf = os.path.abspath(output_pdf)

    if not os.path.exists(abs_docx):
        print(f"[ERRO] Ficheiro não encontrado: {abs_docx}")
        return False

    print(f"[*] A abrir documento no Word: {abs_docx}")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    try:
        doc = word.Documents.Open(abs_docx)

        print("[*] A atualizar Sumário...")
        if doc.TablesOfContents.Count > 0:
            doc.TablesOfContents(1).Update()

        print("[*] A compilar PDF vetorial nativo...")
        doc.ExportAsFixedFormat(
            OutputFileName=abs_pdf,
            ExportFormat=17,  # wdExportFormatPDF
            OpenAfterExport=False,
            OptimizeFor=0,     # wdExportOptimizeForPrint
            CreateBookmarks=1, # wdExportCreateHeadingBookmarks
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )

        doc.Close(SaveChanges=True)
        print(f"[OK] PDF final compilado com sucesso: {abs_pdf}")
        return True

    except Exception as e:
        print(f"[ERRO] Falha durante a compilação: {e}")
        return False
    finally:
        word.Quit()

if __name__ == "__main__":
    compile_pdf("Qualidade_Normalizado.docx", "Qualidade_de_Sistemas_e_Software_Final.pdf")