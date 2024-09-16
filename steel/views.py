from django.shortcuts import render, HttpResponse
import pandas as pd
import os
from django.conf import settings
from django.http import HttpResponse
from django.template import loader


def index(request):
    return render(request,"index.html")

def term(request):
    return HttpResponse("tem pahe hie")

def steel(request):
    static_folder = os.path.join(settings.BASE_DIR, 'static')
    excel_file_path = os.path.join(static_folder, 'steel.xlsx')
    
    sheet_name = 'Database'  
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
    
    def format_dataframe(df):
        pd.set_option('display.float_format', lambda x: f'{x:.2f}')
        return df.to_html(index=False, na_rep='')  

    html_data = format_dataframe(df)

    template = loader.get_template('display_sheet.html')
    context = {
        'table_html': html_data,
    }
    
    return HttpResponse(template.render(context, request))