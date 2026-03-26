#!/usr/bin/env python3
"""
Generate an ODF text document for Avery 5871 business cards.
Sheet: US Letter (8.5" x 11")
Cards: 10 per sheet (2 columns x 5 rows), each 3.5" x 2"
Margins: 0.5" top/bottom, 0.75" left/right (standard Avery 5871)
"""

from odf.opendocument import OpenDocumentText
from odf.style import (Style, TextProperties, TableProperties,
                        TableColumnProperties, TableRowProperties,
                        TableCellProperties, ParagraphProperties,
                        PageLayoutProperties, MasterPage, PageLayout,
                        GraphicProperties)
from odf.table import Table, TableColumn, TableRow, TableCell
from odf.text import P, Span
from odf.draw import Frame, Image

import os

OUTPUT = "/data/git/business-cards/business-cards.odt"

# ── Card Content ──
BUSINESS_NAME = "Lunde Cognitive Effects, Inc."
PERSON_NAME = "Andrew Lunde"
PHONE = "404-416-5688"
EMAIL = "andrew@lunde.com"
LINKEDIN = "linkedin.com/in/andrewlunde"
ADDRESS_1 = "920 Peachtree Battle Avenue NW"
ADDRESS_2 = "Atlanta, GA 30327"

# ── Document Setup ──
doc = OpenDocumentText()

# Page layout: US Letter with Avery 5871 margins
pl = PageLayout(name="PageLayout1")
pl.addElement(PageLayoutProperties(
    pagewidth="8.5in",
    pageheight="11in",
    margintop="0.5in",
    marginbottom="0.5in",
    marginleft="0.75in",
    marginright="0.75in",
))
doc.automaticstyles.addElement(pl)

mp = MasterPage(name="Standard", pagelayoutname=pl)
doc.masterstyles.addElement(mp)

# ── Styles ──

# Table style - fixed width, no borders on the table itself
table_style = Style(name="CardTable", family="table")
table_style.addElement(TableProperties(
    width="7.0in",
    align="center",
))
doc.automaticstyles.addElement(table_style)

# Column style - each column is 3.5"
col_style = Style(name="CardCol", family="table-column")
col_style.addElement(TableColumnProperties(columnwidth="3.5in"))
doc.automaticstyles.addElement(col_style)

# Row style - each row is 2"
row_style = Style(name="CardRow", family="table-row")
row_style.addElement(TableRowProperties(rowheight="2.0in", keeptogether="always"))
doc.automaticstyles.addElement(row_style)

# Cell style - vertical centering, padding, no visible borders
cell_style = Style(name="CardCell", family="table-cell")
cell_style.addElement(TableCellProperties(
    verticalalign="middle",
    paddingtop="0.15in",
    paddingbottom="0.15in",
    paddingleft="0.2in",
    paddingright="0.2in",
    border="none",
))
doc.automaticstyles.addElement(cell_style)

# Business name style - bold, larger
biz_name_style = Style(name="BizName", family="paragraph")
biz_name_style.addElement(ParagraphProperties(
    marginbottom="0.08in",
    textalign="left",
))
biz_name_style.addElement(TextProperties(
    fontsize="11pt",
    fontweight="bold",
    fontfamily="Liberation Sans",
    color="#1a1a1a",
))
doc.automaticstyles.addElement(biz_name_style)

# Person name style
person_style = Style(name="PersonName", family="paragraph")
person_style.addElement(ParagraphProperties(
    marginbottom="0.06in",
    textalign="left",
))
person_style.addElement(TextProperties(
    fontsize="10pt",
    fontweight="bold",
    fontfamily="Liberation Sans",
    color="#333333",
))
doc.automaticstyles.addElement(person_style)

# Contact info style
contact_style = Style(name="ContactInfo", family="paragraph")
contact_style.addElement(ParagraphProperties(
    marginbottom="0.02in",
    textalign="left",
))
contact_style.addElement(TextProperties(
    fontsize="8pt",
    fontfamily="Liberation Sans",
    color="#444444",
))
doc.automaticstyles.addElement(contact_style)

# Separator line style (thin line between name and contact)
sep_style = Style(name="Separator", family="paragraph")
sep_style.addElement(ParagraphProperties(
    marginbottom="0.06in",
    margintop="0.02in",
    textalign="left",
    borderbottom="0.5pt solid #999999",
    paddingbottom="0.04in",
))
sep_style.addElement(TextProperties(fontsize="1pt"))
doc.automaticstyles.addElement(sep_style)


def make_card_content():
    """Generate the content paragraphs for one business card."""
    elements = []

    # Business name
    p1 = P(stylename=biz_name_style, text=BUSINESS_NAME)
    elements.append(p1)

    # Person name
    p2 = P(stylename=person_style, text=PERSON_NAME)
    elements.append(p2)

    # Thin separator
    p_sep = P(stylename=sep_style, text="")
    elements.append(p_sep)

    # Phone
    p3 = P(stylename=contact_style, text=PHONE)
    elements.append(p3)

    # Email
    p4 = P(stylename=contact_style, text=EMAIL)
    elements.append(p4)

    # LinkedIn (without https://www. to save space)
    p5 = P(stylename=contact_style, text=LINKEDIN)
    elements.append(p5)

    # Address
    p6 = P(stylename=contact_style, text=ADDRESS_1)
    elements.append(p6)

    p7 = P(stylename=contact_style, text=ADDRESS_2)
    elements.append(p7)

    return elements


# ── Build the table (2 columns x 5 rows = 10 cards) ──
table = Table(name="BusinessCards", stylename=table_style)
table.addElement(TableColumn(stylename=col_style))
table.addElement(TableColumn(stylename=col_style))

for row_idx in range(5):
    row = TableRow(stylename=row_style)
    for col_idx in range(2):
        cell = TableCell(stylename=cell_style)
        for element in make_card_content():
            cell.addElement(element)
        row.addElement(cell)
    table.addElement(row)

doc.text.addElement(table)

# ── Save ──
doc.save(OUTPUT)
print(f"Created: {OUTPUT}")
print(f"  Sheet: US Letter (8.5\" x 11\")")
print(f"  Cards: 10 (2x5), each 3.5\" x 2\"")
print(f"  Template: Avery 5871")
