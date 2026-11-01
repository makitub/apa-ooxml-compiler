"""
Módulo: apa_docx_engine.py (Versão Definitiva 100% APA 7 v9.0)
Autor: Aristófanes Paulo Gonfulo
Descrição: Normalizador de alto nível para nós OOXML (ECMA-376).
- Transfere o negrito dos títulos para o nível de estilo, eliminando o negrito
  indesejado nos Níveis 2 e 3 do Sumário (TOC).
- Uniformiza categoricamente os 4 níveis de testes com marcador único '• '.
- Formata a norma ISO/IEC e referências com itálico e links a preto (#000000).
- Exibe Relatório de Auditoria detalhado.
"""

import os
import re
import time
import unicodedata
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

class ApaAuditTrail:
    def __init__(self):
        self.logs = []

    def record(self, local, regra, descricao, justificativa_apa):
        self.logs.append({
            "local": local,
            "regra": regra,
            "descricao": descricao,
            "justificativa": justificativa_apa
        })

    def print_report(self):
        print("\n" + "="*85)
        print("          RELATÓRIO DE AUDITORIA E CONFORMIDADE APA (7.ª EDIÇÃO)")
        print("="*85)
        for idx, item in enumerate(self.logs, 1):
            print(f"[{idx}] LOCAL: {item['local']}")
            print(f"    REGRA: {item['regra']}")
            print(f"    AÇÃO:  {item['descricao']}")
            print(f"    PORQUÊ (APA 7): {item['justificativa']}")
            print("-" * 85)
        print(f"Total de intervenções registadas: {len(self.logs)}\n")

audit = ApaAuditTrail()

def set_run_format(run, bold=False, italic=False, font_name="Arial", size_pt=11, force_black=True):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.underline = False
    if force_black:
        run.font.color.rgb = RGBColor(0, 0, 0)

    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)
    rPr.append(rFonts)

    b_el = rPr.find(qn('w:b'))
    if not bold and b_el is not None:
        rPr.remove(b_el)
    elif bold and b_el is None:
        rPr.append(OxmlElement('w:b'))

    i_el = rPr.find(qn('w:i'))
    if not italic and i_el is not None:
        rPr.remove(i_el)
    elif italic and i_el is None:
        rPr.append(OxmlElement('w:i'))

    if force_black:
        color_el = rPr.find(qn('w:color'))
        if color_el is not None:
            rPr.remove(color_el)
        new_color = OxmlElement('w:color')
        new_color.set(qn('w:val'), '000000')
        rPr.append(new_color)

def apply_hanging_indent(paragraph, preview):
    p_pr = paragraph._element.get_or_add_pPr()
    old_ind = p_pr.find(qn('w:ind'))
    if old_ind is not None:
        p_pr.remove(old_ind)
    new_ind = OxmlElement('w:ind')
    new_ind.set(qn('w:left'), '720')
    new_ind.set(qn('w:hanging'), '720')
    p_pr.append(new_ind)

    old_spc = p_pr.find(qn('w:spacing'))
    if old_spc is not None:
        p_pr.remove(old_spc)
    new_spc = OxmlElement('w:spacing')
    new_spc.set(qn('w:line'), '480')
    new_spc.set(qn('w:lineRule'), 'auto')
    new_spc.set(qn('w:before'), '0')
    new_spc.set(qn('w:after'), '0')
    p_pr.append(new_spc)

    audit.record(
        local=f"Referência: {preview}...",
        regra="APA 7 - Secção 9.43 e 2.21",
        descricao="Injeção de recuo francês (1.27 cm / 720 dxa) e espaçamento duplo.",
        justificativa_apa="Primeira linha na margem esquerda e corpo recuado sem linhas vazias extras."
    )

def fix_heading_styles_and_direct_formatting(doc):
    """
    Configura o negrito/itálico diretamente no estilo dos títulos e remove o negrito
    direto dos runs para evitar que o Word propague negrito para o Sumário.
    """
    # 1. Configura os estilos de título
    for style_name, is_bold, is_italic in [
        ('Heading 1', True, False), ('Título 1', True, False),
        ('Heading 2', True, False), ('Título 2', True, False),
        ('Heading 3', True, True),  ('Título 3', True, True)
    ]:
        if style_name in doc.styles:
            doc.styles[style_name].font.bold = is_bold
            doc.styles[style_name].font.italic = is_italic
            doc.styles[style_name].font.name = "Arial"
            doc.styles[style_name].font.size = Pt(11)

    # 2. Configura os estilos de sumário (TOC)
    for style_name, is_bold, is_italic in [
        ('toc 1', True, False),  ('Sumário 1', True, False), ('Sumrio1', True, False),
        ('toc 2', False, False), ('Sumário 2', False, False), ('Sumrio2', False, False),
        ('toc 3', False, True),  ('Sumário 3', False, True),  ('Sumrio3', False, True)
    ]:
        if style_name in doc.styles:
            doc.styles[style_name].font.bold = is_bold
            doc.styles[style_name].font.italic = is_italic
            doc.styles[style_name].font.name = "Arial"
            doc.styles[style_name].font.size = Pt(11)

    # 3. Remove a formatação direta dos runs dos títulos de Nível 2 e 3 no corpo
    for p in doc.paragraphs:
        s_name = p.style.name.lower()
        if 'heading 2' in s_name or 'título 2' in s_name or 'ttulo2' in s_name:
            for r in p.runs:
                rPr = r._r.get_or_add_rPr()
                b_el = rPr.find(qn('w:b'))
                if b_el is not None:
                    rPr.remove(b_el)
            audit.record(
                local=f"Título Nível 2: {p.text[:30]}...",
                regra="APA 7 - Formatação Estrutural",
                descricao="Transferência do negrito para o estilo (remoção de formatação direta no run).",
                justificativa_apa="Garante que o título permanece a negrito no corpo sem poluir o Sumário."
            )
        elif 'heading 3' in s_name or 'título 3' in s_name or 'ttulo3' in s_name:
            for r in p.runs:
                rPr = r._r.get_or_add_rPr()
                b_el = rPr.find(qn('w:b'))
                if b_el is not None:
                    rPr.remove(b_el)
            audit.record(
                local=f"Título Nível 3: {p.text[:30]}...",
                regra="APA 7 - Formatação Estrutural",
                descricao="Transferência de negrito/itálico para o estilo.",
                justificativa_apa="Garante que a entrada no sumário permanece em itálico sem negrito."
            )

def format_references(paragraph):
    txt = unicodedata.normalize('NFC', paragraph.text.replace('\xa0', ' ')).strip()
    if not txt:
        return
    apply_hanging_indent(paragraph, txt[:30])

    # 1. Norma ISO/IEC 25010
    if "iso/iec" in txt.lower() and "25010" in txt:
        canonical_title = "Systems and software engineering — Systems and software Quality Requirements and Evaluation (SQuaRE) — System and software quality models"
        paragraph.text = ""
        r1 = paragraph.add_run("ISO/IEC. (2011). ")
        set_run_format(r1, bold=False, italic=False)
        r2 = paragraph.add_run(canonical_title)
        set_run_format(r2, bold=False, italic=True)
        r3 = paragraph.add_run(" (ISO/IEC Standard No. 25010:2011). https://www.iso.org/standard/35733.html")
        set_run_format(r3, bold=False, italic=False)

        audit.record(
            local="Referência ISO/IEC 25010:2011",
            regra="APA 7 - Secção 9.35 e 9.36",
            descricao="Aplicação de itálico no título oficial e conversão de URL para cor preta neutra.",
            justificativa_apa="Normas técnicas são obras autónomas (título em itálico); links devem ser monocromáticos."
        )
        return

    # 2. Livros clássicos
    books = [
        "Managing quality: The strategic and competitive edge",
        "The art of software testing",
        "Engenharia de software: Uma abordagem profissional",
        "Engenharia de software"
    ]
    for b in books:
        if b.lower() in txt.lower():
            start = txt.lower().find(b.lower())
            end = start + len(b)
            pfx, title, sfx = txt[:start], txt[start:end], txt[end:]
            paragraph.text = ""
            r1 = paragraph.add_run(pfx)
            set_run_format(r1, italic=False)
            r2 = paragraph.add_run(title)
            set_run_format(r2, italic=True)
            if sfx:
                r3 = paragraph.add_run(sfx)
                set_run_format(r3, italic=False)
            audit.record(
                local=f"Referência Livro: {b}",
                regra="APA 7 - Secção 9.19",
                descricao=f"Aplicação de itálico no título '{b}'.",
                justificativa_apa="Livros recebem itálico apenas no título da obra."
            )
            return

    for run in paragraph.runs:
        set_run_format(run, italic=False)

def process_document(docx_path, output_path):
    print(f"[*] A carregar documento: {os.path.abspath(docx_path)}")
    doc = docx.Document(docx_path)

    # 1. Normaliza estilos e desvincula negrito direto dos títulos
    fix_heading_styles_and_direct_formatting(doc)

    in_ref = False
    test_labels = [
        ("testes unitários", "Testes Unitários:"),
        ("testes de integração", "Testes de Integração:"),
        ("testes de sistema", "Testes de Sistema:"),
        ("testes de aceitação", "Testes de Aceitação:")
    ]

    for p in doc.paragraphs:
        raw = unicodedata.normalize('NFC', p.text.replace('\xa0', ' ')).strip()

        # 2. Formata todos os 4 testes da Página 9
        for key_phrase, label in test_labels:
            if key_phrase in raw.lower() and ("operam no nível" in raw.lower() or "avaliam a comunicação" in raw.lower() or "avaliam o ecossistema" in raw.lower() or "representam a etapa" in raw.lower()):
                # Extrai a descrição após o rótulo
                body_match = re.search(rf"{key_phrase}\s*[:]?\s*(.*)", raw, re.IGNORECASE)
                body = body_match.group(1).strip() if body_match else ""

                p_pr = p._element.get_or_add_pPr()
                num_pr = p_pr.find(qn('w:numPr'))
                if num_pr is not None:
                    p_pr.remove(num_pr)

                p.text = ""
                r_b = p.add_run("• ")
                set_run_format(r_b, bold=False)
                r_l = p.add_run(f"{label} ")
                set_run_format(r_l, bold=True)
                r_t = p.add_run(body)
                set_run_format(r_t, bold=False, italic=False)

                old_ind = p_pr.find(qn('w:ind'))
                if old_ind is not None:
                    p_pr.remove(old_ind)
                new_ind = OxmlElement('w:ind')
                new_ind.set(qn('w:left'), '720')
                new_ind.set(qn('w:hanging'), '360')
                p_pr.append(new_ind)

                audit.record(
                    local=f"Lista de Testes: {label}",
                    regra="APA 7 - Formatação Uniforme de Listas",
                    descricao="Injeção de marcador único '• ' e rótulo a negrito.",
                    justificativa_apa="Garante apresentação visual idêntica nos 4 níveis de teste."
                )
                break

        # 3. Referências
        if re.match(r"^refer[êe]ncias[\.\:]?$", raw, re.IGNORECASE):
            in_ref = True
            p.text = "Referências"
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_pr = p._element.get_or_add_pPr()
            if p_pr.find(qn('w:pageBreakBefore')) is None:
                p_pr.append(OxmlElement('w:pageBreakBefore'))
            for run in p.runs:
                set_run_format(run, bold=True)
            continue

        if in_ref and raw:
            format_references(p)

    while doc.paragraphs and not doc.paragraphs[-1].text.strip():
        p_el = doc.paragraphs[-1]._element
        p_el.getparent().remove(p_el)

    doc.save(output_path)
    audit.print_report()
    print(f"[OK] Ficheiro normalizado gerado em: {output_path}")

if __name__ == "__main__":
    process_document("Qualidade de Sistemas e Software.docx", "Qualidade_Normalizado.docx")