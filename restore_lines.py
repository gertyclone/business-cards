#!/usr/bin/env python3
"""
Restore the red design lines from the original business-cards-lcfx.odt
into the current replicated version.
"""

import zipfile
import io
from xml.dom.minidom import parseString

ORIGINAL = "/tmp/original-lcfx.odt"
CURRENT = "/data/git/business-cards/business-cards-lcfx.odt"
OUTPUT = "/data/git/business-cards/business-cards-lcfx.odt"

# Read original to extract the lines and their context
with zipfile.ZipFile(ORIGINAL, 'r') as z:
    orig_content = z.read('content.xml').decode('utf-8')

orig_dom = parseString(orig_content)

# Find the outer BusinessCards table lines from original
# These are the 5 lines anchored in BusinessCards cells
orig_lines_info = []
for line in orig_dom.getElementsByTagName('draw:line'):
    name = line.getAttribute('draw:name')
    parent = line.parentNode
    while parent:
        if hasattr(parent, 'tagName') and parent.tagName == 'table:table':
            if parent.getAttribute('table:name') == 'BusinessCards':
                # Get paragraph and cell context
                p_elem = line.parentNode  # the text:p containing this line
                cell = p_elem.parentNode
                row = cell.parentNode
                
                # Find row and cell indices
                outer = parent
                row_idx = 0
                for r in outer.childNodes:
                    if r == row:
                        break
                    if r.nodeType == r.ELEMENT_NODE and r.tagName == 'table:table-row':
                        row_idx += 1
                
                cell_idx = 0
                for c in row.childNodes:
                    if c == cell:
                        break
                    if c.nodeType == c.ELEMENT_NODE and c.tagName == 'table:table-cell':
                        cell_idx += 1
                
                orig_lines_info.append({
                    'name': name,
                    'line_xml': line.toxml(),
                    'p_style': p_elem.getAttribute('text:style-name'),
                    'row': row_idx,
                    'col': cell_idx,
                })
            break
        parent = parent.parentNode

print(f"Found {len(orig_lines_info)} lines to restore:")
for info in orig_lines_info:
    print(f"  {info['name']} at row {info['row']}, col {info['col']} (p style: {info['p_style']})")

# Now read current file and inject the lines
with zipfile.ZipFile(CURRENT, 'r') as z:
    files = {}
    for name in z.namelist():
        files[name] = z.read(name)

cur_dom = parseString(files['content.xml'].decode('utf-8'))

# Find current BusinessCards table
outer = None
for t in cur_dom.getElementsByTagName('table:table'):
    if t.getAttribute('table:name') == 'BusinessCards':
        outer = t
        break

# Get direct child rows
direct_rows = [r for r in outer.childNodes 
               if r.nodeType == r.ELEMENT_NODE and r.tagName == 'table:table-row']

# Check if P3, P9 styles exist; they should from the original styles preserved in the file
# Add lines to the appropriate cells
line_counter = 0
for info in orig_lines_info:
    row_idx = info['row']
    col_idx = info['col']
    
    # Clamp row index to available rows (we have 5 rows: 0-4)
    if row_idx >= len(direct_rows):
        row_idx = len(direct_rows) - 1
    
    row = direct_rows[row_idx]
    cells = [c for c in row.childNodes 
             if c.nodeType == c.ELEMENT_NODE and c.tagName == 'table:table-cell']
    
    if col_idx >= len(cells):
        col_idx = len(cells) - 1
    
    cell = cells[col_idx]
    
    # Create a paragraph with the line
    line_counter += 1
    p = cur_dom.createElement('text:p')
    p.setAttribute('text:style-name', info['p_style'])
    
    # Parse the line XML and import it
    line_frag = parseString(f'<root xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">{info["line_xml"]}</root>')
    line_elem = line_frag.documentElement.firstChild
    imported = cur_dom.importNode(line_elem, True)
    
    # Give unique name
    imported.setAttribute('draw:name', f'{info["name"]}_restored{line_counter}')
    
    p.appendChild(imported)
    
    # Insert at the beginning of the cell (before the inner table)
    if cell.firstChild:
        cell.insertBefore(p, cell.firstChild)
    else:
        cell.appendChild(p)
    
    print(f"  Restored {info['name']} -> row {row_idx}, col {col_idx}")

# Also ensure gr5 and gr6 styles exist (they were in the original)
# Check if they're already in the automatic styles
auto_styles = cur_dom.getElementsByTagName('office:automatic-styles')[0]
existing_styles = set()
for s in auto_styles.getElementsByTagName('style:style'):
    existing_styles.add(s.getAttribute('style:name'))

# Get gr5, gr6, P3, P9, P25, P26 from original if missing
orig_auto = orig_dom.getElementsByTagName('office:automatic-styles')[0]
needed = {'gr5', 'gr6', 'P3', 'P9', 'P25', 'P26'}
for s in orig_auto.getElementsByTagName('style:style'):
    name = s.getAttribute('style:name')
    if name in needed and name not in existing_styles:
        imported = cur_dom.importNode(s, True)
        auto_styles.appendChild(imported)
        print(f"  Added missing style: {name}")

# Write back
content_out = cur_dom.toxml()
buf = io.BytesIO()
with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
    for name, data in files.items():
        if name == 'content.xml':
            zout.writestr(name, content_out.encode('utf-8'))
        else:
            zout.writestr(name, data)

with open(OUTPUT, 'wb') as f:
    f.write(buf.getvalue())

print(f"\nDone: {OUTPUT}")
