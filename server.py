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

WORKSPACE_DIR = BASE_DIR
if not os.path.exists(os.path.join(WORKSPACE_DIR, "inventario_max_technology_store.xlsx")):
    WORKSPACE_DIR = "d:/Documents/CLIENTES/MAX TECHNOLOGY"

EXCEL_PATH = os.path.join(WORKSPACE_DIR, "inventario_max_technology_store.xlsx")
JSON_PATH = os.path.join(WORKSPACE_DIR, "master_consolidated_inventory.json")
IMG_DIR = os.path.join(WORKSPACE_DIR, "IMAGENES DE PRODUCTOS")
HTML_INDEX_PATH = os.path.join(WORKSPACE_DIR, "index.html")
HTML_CATALOG_PATH = os.path.join(WORKSPACE_DIR, "catalogo_max_tech.html")

def rebuild_static_html(products):
    items_json_str = json.dumps(products, ensure_ascii=False)
    
    # Check if index.html exists, replace products in it
    if os.path.exists(HTML_INDEX_PATH):
        with open(HTML_INDEX_PATH, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Replace defaultProducts
        marker_start = 'const defaultProducts = '
        marker_end = ';\n\n    // Load custom stock overrides'
        if marker_start in content and marker_end in content:
            before = content.split(marker_start)[0] + marker_start
            after = marker_end + content.split(marker_end)[1]
            new_content = before + items_json_str + after
            with open(HTML_INDEX_PATH, 'w', encoding='utf-8') as f_out:
                f_out.write(new_content)
            with open(HTML_CATALOG_PATH, 'w', encoding='utf-8') as f_out:
                f_out.write(new_content)
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

                products = save_products_json(products, JSON_PATH)
                if os.path.exists(EXCEL_PATH):
                    sync_json_to_excel(products, EXCEL_PATH)
                rebuild_static_html(products)

                self.send_json({
                    'success': True,
                    'message': '¡Producto guardado y sincronizado!',
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
                    'message': f'Producto {sku} eliminado.',
                    'products': products
                })
            except Exception as e:
                self.send_json({'success': False, 'error': str(e)}, 500)
            return

        if path == '/api/products/upload-image':
            try:
                filename = req_data.get('filename', 'producto_nuevo.jpg')
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
                # 1. Update data & rebuild HTML
                products = load_products_json(JSON_PATH)
                rebuild_static_html(products)
                if os.path.exists(EXCEL_PATH):
                    sync_json_to_excel(products, EXCEL_PATH)

                # 2. Git operations
                subprocess.run(["git", "add", "."], cwd=WORKSPACE_DIR, capture_output=True, text=True)
                commit_msg = req_data.get('message') or "Actualización de catálogo, inventario y precios"
                p_commit = subprocess.run(["git", "commit", "-m", commit_msg], cwd=WORKSPACE_DIR, capture_output=True, text=True)
                p_push = subprocess.run(["git", "push", "origin", "main"], cwd=WORKSPACE_DIR, capture_output=True, text=True)

                output_log = (p_push.stdout or "") + (p_push.stderr or "")
                
                self.send_json({
                    'success': True,
                    'message': '¡Catálogo publicado en GitHub Pages con éxito!',
                    'log': output_log,
                    'url': 'https://andreasilector-sketch.github.io/max-technology-inventario/'
                })
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
