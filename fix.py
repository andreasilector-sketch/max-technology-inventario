import re

with open('max_inventory_core.py', 'r', encoding='utf-8') as f:
    content = f.read()

new_sync = '''def sync_json_to_excel(products, excel_path):
    try:
        wb = openpyxl.load_workbook(excel_path)
    except Exception as e:
        return False
    ws = wb['🛒 Inventario Maestro']
    
    ranges_to_remove = []
    for rng in ws.merged_cells.ranges:
        if rng.max_row >= 4:
            ranges_to_remove.append(rng)
    for rng in ranges_to_remove:
        ws.unmerge_cells(str(rng))
        
    max_r = max(ws.max_row, len(products) + 15)
    for r in range(4, max_r + 1):
        for c in range(1, 25):
            safe_set_cell(ws, r, c, None)'''

old_sync = '''def sync_json_to_excel(products, excel_path):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb['🛒 Inventario Maestro']
    
    # Clear existing data rows starting from row 4
    max_r = max(ws.max_row, 100)
    for r in range(4, max_r + 1):
        for c in range(1, 25):
            ws.cell(row=r, column=c).value = None'''

content = content.replace(old_sync, new_sync)

content = re.sub(r'ws\.cell\(\s*row=([a-zA-Z0-9_]+),\s*column=([a-zA-Z0-9_]+)\s*\)', r'safe_set_cell(ws, \1, \2)', content)
content = re.sub(r'ws\.cell\(\s*row=([a-zA-Z0-9_]+),\s*column=([a-zA-Z0-9_]+),\s*value=(.+?)\)', r'safe_set_cell(ws, \1, \2, \3)', content)

with open('max_inventory_core.py', 'w', encoding='utf-8') as f:
    f.write(content)