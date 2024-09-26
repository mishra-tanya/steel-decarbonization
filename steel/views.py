from django.shortcuts import render, HttpResponse
import pandas as pd
import os
import io
from django.conf import settings
from django.http import HttpResponse
from django.template import loader
import matplotlib.pyplot as plt
from io import BytesIO
import matplotlib.colors as mcolors
import pythoncom
import win32com.client as win32
from django.conf import settings
from django.shortcuts import render
from openpyxl import load_workbook
import xlwings as xw
import openpyxl
from PIL import Image
from PIL import ImageGrab
import time

def index(request):
    return render(request,"index.html")

def term(request):
    return render(request, 'starter-page.html')

def ste(request):
    return render(request,"steel.html")

def steel(request):
    # Define the path to the Excel file
    static_folder = os.path.join(settings.BASE_DIR, 'static')
    excel_file_path = os.path.join(static_folder, 'steel.xlsx')

    if not os.path.exists(excel_file_path):
        return render(request, 'index.html', {'error': 'Excel file not found.'})

    # Ensure the file is not open elsewhere and we can access it
    try:
        wb = openpyxl.load_workbook(excel_file_path)
        app = xw.App(visible=False)  # Hide Excel instance
        wb = app.books.open(excel_file_path)
        sheet = wb.sheets['ironandsteelmakertool']

        # Update Excel cells with user inputs
        sheet.range('D16').value = request.GET.get('start_year')
        sheet.range('D17').value = request.GET.get('base_year_activity')
        sheet.range('D18').value = request.GET.get('base_year_emission')
        sheet.range('D20').value = request.GET.get('end_year')
        sheet.range('D21').value = request.GET.get('target_activity')
        sheet.range('D22').value = request.GET.get('target_year_output')
        sheet.range('D24').value = request.GET.get('scrap_base')
        sheet.range('D25').value = request.GET.get('scrap_target')
        time.sleep(2) 
        save_directory = os.path.join(settings.BASE_DIR, 'static/img')
        os.makedirs(save_directory, exist_ok=True) 

        start_row = 30
        end_row = 100
        start_col = 'A'
        end_col = 'AD'

      
        last_row = min(end_row, sheet.cells.last_cell.row)

        capture_range = sheet.range(f'{start_col}{start_row}:{end_col}{last_row}')

        capture_range.copy()

        time.sleep(1)

        screenshot = ImageGrab.grabclipboard()

        image_name = 'steel_content_image.png'
        server_image_path = os.path.join(save_directory, image_name)

        if screenshot is not None:
            screenshot.save(server_image_path)
            print(f"Content screenshot successfully saved to {server_image_path}")
        else:
            print("Failed to grab the image from clipboard.")

        wb.close()
        app.quit()
        return render(request, 'index.html', {'error': str(e)})

    except Exception as e:
        if 'wb' in locals():
            wb.close()  # Ensure workbook is closed if an error occurs
        if 'app' in locals():
            app.quit()  # Ensure Excel app is closed if an error occurs
        return render(request, 'index.html', {'error': str(e)})


def format_dataframe(df):
    pd.set_option('display.float_format', lambda x: f'{x:.2f}')
    return df.to_html(index=False, na_rep='')