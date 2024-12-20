from http.server import BaseHTTPRequestHandler
from os.path import dirname, abspath, join
dir = dirname(abspath(__file__))
 


#csvs_to_excel_with_formatting(output_csv_files,'NekilHolkin.xlsx')


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-type', 'text/plain')
        self.end_headers()
        self.wfile.write(str('Hello Mundo!!').encode())
        return