def create_slip(doc, doc_type, batch_id, number):
    section = doc.sections[0]

    # Create FULL-PAGE table
    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    table.allow_autofit = False

    table.columns[0].width = section.page_width - section.left_margin - section.right_margin
    table.rows[0].height = section.page_height - section.top_margin - section.bottom_margin

    set_table_border(table)

    cell = table.cell(0, 0)

    # Add padding inside box
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for side in ('top', 'left', 'bottom', 'right'):
        mar = OxmlElement(f'w:{side}')
        mar.set(qn('w:w'), '800')  # padding
        mar.set(qn('w:type'), 'dxa')
        tcMar.append(mar)
    tcPr.append(tcMar)

    cell.paragraphs[0].clear()

    def add_line(text, bold=False, size=22):
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)

    # ---- CONTENT ----
    if doc_type == "MOD":
        add_line("F074025-000000", size=24)
        add_line("GUAR GUM POWDER", size=26)
        add_line("MODIFIED", size=26)
        add_line("")
    else:
        add_line("FARINA GUAR 200 MESH 5000 T/C", size=28)
        add_line("")

    add_line("NET WEIGHT: 900 KG", size=24)
    add_line("GROSS WEIGHT: 903 KG", size=24)
    add_line("(Without Pallet)", size=20)
    add_line("")
    add_line(f"BATCH NO.: {batch_id} ({number})", bold=True, size=30)
