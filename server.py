import os
import sys
import json
import math
import mimetypes
import subprocess
import urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler
import base64

# Import core functions
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.append(BASE_DIR)

from max_inventory_core import (
    calculate_prices,
    generate_wa_message,
    load_products_json,
    save_products_json,
    sync_json_to_excel,
    sync_excel_to_json,
    roundup_price
)

WORKSPACE_DIR = os.path.abspath(os.path.join(BASE_DIR, "../../../../.."))
# Fallback to current directory if not matching
if not os.path.exists(os.path.join(WORKSPACE_DIR, "inventario_max_technology_store.xlsx")):
    WORKSPACE_DIR = "d:/Documents/CLIENTES/MAX TECHNOLOGY"

EXCEL_PATH = os.path.join(WORKSPACE_DIR, "inventario_max_technology_store.xlsx")
JSON_PATH = os.path.join(BASE_DIR, "master_consolidated_inventory.json")
IMG_DIR = os.path.join(WORKSPACE_DIR, "IMAGENES DE PRODUCTOS")
HTML_INDEX_PATH = os.path.join(WORKSPACE_DIR, "index.html")
HTML_CATALOG_PATH = os.path.join(WORKSPACE_DIR, "catalogo_max_tech.html")

def rebuild_static_html(products):
    items_json_str = json.dumps(products, ensure_ascii=False)
    
    # We load the template from our catalog generator
    template_path = os.path.join(BASE_DIR, "build_interactive_catalog.py")
    if os.path.exists(template_path):
        with open(template_path, 'r', encoding='utf-8') as f:
            code = f.read()
            # Extract html_content format string
            start_marker = 'html_content = f"""'
            end_marker = '"""\n\n# Write to catalogo_max_tech.html'
            if start_marker in code and end_marker in code:
                tmpl = code.split(start_marker)[1].split(end_marker)[0]
                html_rendered = tmpl.replace('{items_json_str}', items_json_str)
                with open(HTML_INDEX_PATH, 'w', encoding='utf-8') as f_out:
                    f_out.write(html_rendered)
                with open(HTML_CATALOG_PATH, 'w', encoding='utf-8') as f_out:
                    f_out.write(html_rendered)
                return True
    return False

class MaxTechHandler(BaseHTTPRequestHandler):
    def send_json(self, data, status=200):
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path in ['/', '/admin', '/admin.html']:
            admin_file = os.path.join(WORKSPACE_DIR, 'admin.html')
            if os.path.exists(admin_file):
                self.serve_file(admin_file, 'text/html; charset=utf-8')
            else:
                self.serve_file(os.path.join(WORKSPACE_DIR, 'index.html'), 'text/html; charset=utf-8')
            return

        if path == '/api/products':
            products = load_products_json(JSON_PATH)
            if not products and os.path.exists(EXCEL_PATH):
                products = sync_excel_to_json(EXCEL_PATH, JSON_PATH)
            self.send_json({'success': True, 'products': products})
            return

        # Serve static files from workspace directory
        clean_path = path.lstrip('/')
        clean_path = urllib.parse.unquote(clean_path)
        full_path = os.path.join(WORKSPACE_DIR, clean_path)

        if os.path.exists(full_path) and os.path.isfile(full_path):
            mime_type, _ = mimetypes.guess_type(full_path)
            self.serve_file(full_path, mime_type or 'application/octet-stream')
            return

        self.send_error(404, f"Archivo no encontrado: {path}")

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(length) if length > 0 else b''

        try:
            req_data = json.loads(body.decode('utf-8')) if body else {}
        except Exception:
            req_data = {}

        if path == '/api/products/save':
            try:
                products = load_products_json(JSON_PATH)
                item = req_data.get('product')
                if not item:
                    # Batch save
                    products = req_data.get('products', products)
                else:
                    sku = item.get('sku')
                    found = False
                    for i, p in enumerate(products):
                        if p.get('sku') == sku:
                            products[i] = item
                            found = True
                            break
                    if not found:
                        products.append(item)

                # Recalculate & Save
                products = save_products_json(products, JSON_PATH)
                if os.path.exists(EXCEL_PATH):
                    sync_json_to_excel(products, EXCEL_PATH)
                rebuild_static_html(products)

                self.send_json({
                    'success': True,
                    'message': '¡Producto guardado y sincronizado con Excel y Web!',
                    'products': products
                })
            except Exception as e:
                self.send_json({'success': False, 'error': str(e)}, 500)
            return

        if path == '/api/products/delete':
            try:
                sku = req_data.get('sku')
                products = load_products_json(JSON_PATH)
                products = [p for p in products if p.get('sku') != sku]
                
                products = save_products_json(products, JSON_PATH)
                if os.path.exists(EXCEL_PATH):
                    sync_json_to_excel(products, EXCEL_PATH)
                rebuild_static_html(products)

                self.send_json({
                    'success': True,
                    'message': f'Producto {sku} eliminado correctamente.',
                    'products': products
                })
            except Exception as e:
                self.send_json({'success': False, 'error': str(e)}, 500)
            return

        if path == '/api/products/upload-image':
            try:
                filename = req_data.get('filename', 'producto_nuevo.jpg')
                # clean filename
                filename = os.path.basename(filename).replace(' ', '_')
                b64data = req_data.get('base64', '')
                if ',' in b64data:
                    b64data = b64data.split(',')[1]
                
                os.makedirs(IMG_DIR, exist_ok=True)
                target_file = os.path.join(IMG_DIR, filename)
                with open(target_file, 'wb') as f:
                    f.write(base64.b64decode(b64data))

                rel_path = f"IMAGENES DE PRODUCTOS/{filename}"
                self.send_json({'success': True, 'path': rel_path})
            except Exception as e:
                self.send_json({'success': False, 'error': str(e)}, 500)
            return

        if path == '/api/sync-from-excel':
            try:
                products = sync_excel_to_json(EXCEL_PATH, JSON_PATH)
                rebuild_static_html(products)
                self.send_json({
                    'success': True,
                    'count': len(products),
                    'message': f'Se sincronizaron {len(products)} referencias desde el Excel Maestro.',
                    'products': products
                })
            except Exception as e:
                self.send_json({'success': False, 'error': str(e)}, 500)
            return

        if path == '/api/open-excel':
            try:
                if os.path.exists(EXCEL_PATH):
                    os.startfile(EXCEL_PATH)
                    self.send_json({'success': True, 'message': 'Excel abierto en Windows.'})
                else:
                    self.send_json({'success': False, 'error': 'Archivo Excel no encontrado.'}, 404)
            except Exception as e:
                self.send_json({'success': False, 'error': str(e)}, 500)
            return

        if path == '/api/git-sync':
            try:
                # 1. Make sure static files are fresh
                products = load_products_json(JSON_PATH)
                rebuild_static_html(products)
                if os.path.exists(EXCEL_PATH):
                    sync_json_to_excel(products, EXCEL_PATH)

                # 2. Run Git commands
                cmd_add = ["git", "add", "."]
                p_add = subprocess.run(cmd_add, cwd=WORKSPACE_DIR, capture_output=True, text=True)

                commit_msg = req_data.get('message') or "Actualización de catálogo, inventario y precios"
                cmd_commit = ["git", "commit", "-m", commit_msg]
                p_commit = subprocess.run(cmd_commit, cwd=WORKSPACE_DIR, capture_output=True, text=True)

                cmd_push = ["git", "push", "origin", "main"]
                p_push = subprocess.run(cmd_push, cwd=WORKSPACE_DIR, capture_output=True, text=True)

                output_log = (
                    f"Git Add: {p_add.stdout or 'OK'}\n"
                    f"Git Commit: {p_commit.stdout or p_commit.stderr or 'No changes'}\n"
                    f"Git Push: {p_push.stdout or p_push.stderr}"
                )

                if p_push.returncode == 0 or "Everything up-to-date" in output_log or "up to date" in output_log.lower():
                    self.send_json({
                        'success': True,
                        'message': '¡Catálogo publicado en GitHub Pages con éxito!',
                        'log': output_log,
                        'url': 'https://andreasilector-sketch.github.io/max-technology-inventario/'
                    })
                else:
                    self.send_json({
                        'success': False,
                        'message': 'Error al hacer push a GitHub.',
                        'log': output_log
                    }, 500)
            except Exception as e:
                self.send_json({'success': False, 'error': str(e)}, 500)
            return

        self.send_error(404, "Ruta de API no encontrada.")

    def serve_file(self, full_path, content_type):
        try:
            with open(full_path, 'rb') as f:
                content = f.read()
            self.send_response(200)
            self.send_header('Content-Type', content_type)
            self.send_header('Content-Length', str(len(content)))
            self.end_headers()
            self.wfile.write(content)
        except Exception as e:
            self.send_error(500, f"Error al leer archivo: {e}")

def run_server(port=8000):
    server_address = ('127.0.0.1', port)
    try:
        httpd = HTTPServer(server_address, MaxTechHandler)
        print(f"=====================================================")
        print(f"  MAX TECHNOLOGY STORE - SERVIDOR LOCAL ACTIVO")
        print(f"  Panel de Administración: http://localhost:{port}/admin")
        print(f"  Catálogo Público:        http://localhost:{port}/")
        print(f"=====================================================")
        httpd.serve_forever()
    except OSError as e:
        if port < 8010:
            run_server(port + 1)
        else:
            print(f"Error starting server: {e}")

if __name__ == '__main__':
    port = 8000
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            pass
    run_server(port)
