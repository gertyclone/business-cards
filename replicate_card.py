#!/usr/bin/env python3
"""
Take the bottom-right card design (Table3) from business-cards-lcfx.odt
and replicate it into all 10 card positions. Remove the second page.
"""

import zipfile
import copy
import io
import os
import re
from xml.dom.minidom import parseString

SRC = "/data/git/business-cards/business-cards-lcfx.odt"
OUT = "/data/git/business-cards/business-cards-lcfx.odt"

# Read the ODF zip
with zipfile.ZipFile(SRC, 'r') as zin:
    files = {}
    for name in zin.namelist():
        files[name] = zin.read(name)

content_xml = files['content.xml'].decode('utf-8')
dom = parseString(content_xml)

# Find the outer BusinessCards table
outer_table = None
for t in dom.getElementsByTagName('table:table'):
    if t.getAttribute('table:name') == 'BusinessCards':
        outer_table = t
        break

if not outer_table:
    raise Exception("BusinessCards table not found")

# Find Table3 (the bottom-right card design)
table3 = None
for t in dom.getElementsByTagName('table:table'):
    if t.getAttribute('table:name') == 'Table3':
        table3 = t
        break

if not table3:
    raise Exception("Table3 not found")

# Get the cell that contains Table3 (its parent)
table3_cell = table3.parentNode

# Build a template card cell: clone Table3's parent cell
# But we need to handle the draw shapes having unique names
shape_counter = [0]
def make_card_cell(cell_style):
    """Create a new cell with a copy of the Table3 card design."""
    global shape_counter
    
    # Clone the entire cell content from table3_cell
    new_cell = table3_cell.cloneNode(True)
    new_cell.setAttribute('table:style-name', cell_style)
    
    # Rename the inner table to avoid conflicts
    shape_counter[0] += 1
    suffix = shape_counter[0]
    
    for t in new_cell.getElementsByTagName('table:table'):
        old_name = t.getAttribute('table:name')
        t.setAttribute('table:name', f'{old_name}_copy{suffix}')
    
    # Rename draw shapes to be unique
    for elem in new_cell.getElementsByTagName('draw:custom-shape'):
        old_name = elem.getAttribute('draw:name')
        elem.setAttribute('draw:name', f'{old_name}_c{suffix}')
    for elem in new_cell.getElementsByTagName('draw:line'):
        old_name = elem.getAttribute('draw:name')
        elem.setAttribute('draw:name', f'{old_name}_c{suffix}')
    
    return new_cell

# Get the cell style used by the outer table
cell_style = 'BusinessCards.A1'

# Clear all existing rows from outer table
rows = list(outer_table.getElementsByTagName('table:table-row'))
# Only remove direct child rows (not nested ones)
direct_rows = [r for r in rows if r.parentNode == outer_table]
for row in direct_rows:
    outer_table.removeChild(row)

# Also remove Table1, Table2, Table3 and any other content that's outside
# the outer table (to remove the second page)
body = dom.getElementsByTagName('office:text')[0]

# Collect all direct children of office:text that come after the outer table
# These constitute the "second page" content
nodes_to_remove = []
found_table = False
for child in list(body.childNodes):
    if child == outer_table:
        found_table = True
        continue
    if found_table:
        nodes_to_remove.append(child)

for node in nodes_to_remove:
    body.removeChild(node)

# Also remove any content before the outer table that's not needed
# (keep style declarations etc, just remove extra paragraphs/tables)
nodes_before = []
for child in list(body.childNodes):
    if child == outer_table:
        break
    if child.nodeType == child.ELEMENT_NODE:
        tag = child.tagName
        if tag in ('text:p', 'table:table') and child != outer_table:
            nodes_before.append(child)

for node in nodes_before:
    body.removeChild(node)

# Build 5 rows x 2 columns of the card design
row_style = 'BusinessCards.1'

for row_idx in range(5):
    # Create row element
    row = dom.createElement('table:table-row')
    row.setAttribute('table:style-name', row_style)
    
    for col_idx in range(2):
        cell = make_card_cell(cell_style)
        row.appendChild(cell)
    
    outer_table.appendChild(row)

# Serialize back
content_out = dom.toxml()

# Write back to ODF zip
buf = io.BytesIO()
with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
    for name, data in files.items():
        if name == 'content.xml':
            zout.writestr(name, content_out.encode('utf-8'))
        else:
            zout.writestr(name, data)

with open(OUT, 'wb') as f:
    f.write(buf.getvalue())

print(f"Created: {OUT}")
print("  10 cards (2x5), all matching the bottom-right design")
print("  Second page removed")
