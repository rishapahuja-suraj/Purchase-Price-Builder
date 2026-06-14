from http.server import BaseHTTPRequestHandler
import json, csv, io, re
import urllib.request, urllib.parse

def extract_sheet_id(url_or_id):
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_\-]+)', str(url_or_id))
    return m.group(1) if m else str(url_or_id).strip()

def fetch_sheet_csv(spreadsheet_id):
    url = f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv'
    req = urllib.request.Request(
        url,
        headers={'User-Agent': 'Mozilla/5.0 (compatible; PerformancePlanBuilder/1.0)'}
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        # Follow redirects (Google redirects to login if not shared — catch that)
        final_url = resp.geturl()
        if 'accounts.google.com' in final_url or 'ServiceLogin' in final_url:
            raise ValueError(
                'Sheet is not publicly accessible. '
                'Share it with "Anyone with the link can view" and try again.'
            )
        content = resp.read().decode('utf-8')
    reader = csv.DictReader(io.StringIO(content))
    rows = [dict(row) for row in reader]
    if not rows:
        raise ValueError('Sheet appears to be empty or has no header row.')
    return rows

class handler(BaseHTTPRequestHandler):

    def do_GET(self):
        """Health check."""
        payload = json.dumps({'status': 'ok', 'function': 'gsheet'}).encode()
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(payload)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(payload)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Content-Length', '0')
        self.end_headers()

    def do_POST(self):
        try:
            length = int(self.headers.get('Content-Length') or 0)
            body   = self.rfile.read(length)
            data   = json.loads(body.decode('utf-8'))

            url = data.get('url', '').strip()
            if not url:
                raise ValueError('No URL provided.')

            sheet_id = extract_sheet_id(url)
            if not sheet_id:
                raise ValueError('Could not extract a spreadsheet ID from the URL.')

            rows = fetch_sheet_csv(sheet_id)
            payload = json.dumps({'rows': rows, 'count': len(rows)}).encode()

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(payload)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(payload)

        except Exception as e:
            import traceback
            err = json.dumps({'error': str(e), 'detail': traceback.format_exc()}).encode()
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(err)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(err)

    def log_message(self, *args):
        pass
