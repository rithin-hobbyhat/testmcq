import subprocess, http.server, threading, re, os

class QuietHandler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

httpd = http.server.HTTPServer(('127.0.0.1', 8888), QuietHandler)
t = threading.Thread(target=httpd.serve_forever, daemon=True)
t.start()

edge_path = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
if not os.path.exists(edge_path):
    edge_path = r'C:\Program Files\Microsoft\Edge\Application\msedge.exe'

cmd = [edge_path, '--headless', '--dump-dom', 'http://127.0.0.1:8888/web/index.html']
out = subprocess.check_output(cmd, stderr=subprocess.STDOUT, text=True, encoding='utf-8')

if 'HOBBY HAT MCQ' in out:
    print('PASS: HOBBY HAT MCQ is present in rendered DOM')
else:
    print('FAIL: HOBBY HAT MCQ not found in DOM')

m = re.search(r'<div class="brand-title">([^<]+)</div>', out)
if m:
    print('Brand Title:', m.group(1))

m_sub = re.search(r'<div class="brand-sub"[^>]*>([^<]+)</div>', out)
if m_sub:
    print('Brand Subtitle:', m_sub.group(1))
