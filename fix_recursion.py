import sys
content = open('max_inventory_core.py', 'r', encoding='utf-8').read()
content = content.replace('cell = safe_set_cell(ws, r, c)', 'cell = ws.cell(row=r, column=c)', 1)
open('max_inventory_core.py', 'w', encoding='utf-8').write(content)