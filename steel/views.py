from django.shortcuts import render, HttpResponse
import pandas as pd
import os
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

    # Get user inputs from the GET request
    start_year = request.GET.get('start_year')
    base_year_activity = request.GET.get('base_year_activity')
    base_year_emission = request.GET.get('base_year_emission')
    end_year = request.GET.get('end_year')
    target_activity = request.GET.get('target_activity')
    target_year_output = request.GET.get('target_year_output')
    scrap_base = request.GET.get('scrap_base')
    scrap_target = request.GET.get('scrap_target')

    # Load the Excel file
    workbook = load_workbook(excel_file_path)
    sheet = workbook['ironandsteelmakertool']  # Adjust sheet name accordingly

    # Update Excel cells with user inputs
    sheet['D16'] = start_year
    sheet['D17'] = base_year_activity
    sheet['D18'] = base_year_emission
    sheet['D20'] = end_year
    sheet['D21'] = target_activity
    sheet['D22'] = target_year_output
    sheet['D24'] = scrap_base
    sheet['D25'] = scrap_target

    # Save the workbook after making changes
    workbook.save(excel_file_path)

    # Return success message to the template
    return render(request, 'steel_graph.html', {'message': 'Excel updated successfully!'})


def format_dataframe(df):
    pd.set_option('display.float_format', lambda x: f'{x:.2f}')
    return df.to_html(index=False, na_rep='')