import requests
import pandas as pd
import json
import csv
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
import re


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