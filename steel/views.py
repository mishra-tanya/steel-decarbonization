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
        # Ensure df is a DataFrame, and format it as HTML
        pd.set_option('display.float_format', lambda x: f'{x:.2f}')
        return df.to_html(index=False, na_rep='')

    # Filter rows where 'Scenario' is 'SBTi 1.5C' and 'Sector.ETP' is 'Iron and steel - core boundary'
    specific_row = df[(df['Scenario'] == 'SBTi 1.5C') & (df['Sector.ETP'] == 'Iron and steel - core boundary')]

    # Identify columns that contain years
    year_columns = [str(year) for year in range(2023, 2029)]
    
    # Get all columns and identify the first 8 columns
    all_columns = df.columns
    first_8_columns = list(all_columns[:8])
    
    # Identify non-year columns and combine with year columns
    non_year_columns = [col for col in all_columns if col not in year_columns]
    
    # Combine first 8 columns with year columns (ensure unique columns)
    combined_columns = list(dict.fromkeys(first_8_columns + year_columns))

    # Ensure we have columns to select
    if not combined_columns:
        return HttpResponse("No columns available to display.")
    
    specific_row = specific_row[combined_columns]

    # If specific_row is empty, handle it
    if specific_row.empty:
        return HttpResponse("No data found for 'SBTi 1.5C' scenario and 'Iron and steel - core boundary'.")
    
    # Convert the filtered DataFrame to HTML
    html_data = format_dataframe(specific_row)

    # Load the template
    template = loader.get_template('display_sheet.html')
    context = {
        'table_html': html_data,
    }
    
    return HttpResponse(template.render(context, request))