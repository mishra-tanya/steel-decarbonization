from django.shortcuts import render, redirect
from django.http import JsonResponse
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
import datetime
from .forms import ContactForm
from django.contrib import messages

from django.contrib.auth.models import User
from django.contrib import messages
from django.contrib.auth import login
from .models import Profile 
from django.contrib.auth import authenticate, login as auth_login, logout as auth_logout  # Import necessary functions
from .middlewares import auth
# index
def index(request):
    return render(request,"index.html")

# starter page
def term(request):
    return render(request, 'starter-page.html')

# steel page
@auth
def ste(request):
    return render(request,"steel.html")

# steel logic url steel decarbonisation
@auth
def steel(request):
    # Define the path to the Excel file
    static_folder = os.path.join(settings.BASE_DIR, 'static')
    excel_file_path = os.path.join(static_folder, 'steel.xlsx')

    if not os.path.exists(excel_file_path):
        return render(request, 'index.html', {'error': 'Excel file not found.'})

    # Ensure the file is not open elsewhere and we can access it
    try:
        wb = openpyxl.load_workbook(excel_file_path)
        app = xw.App(visible=False)  
        wb = app.books.open(excel_file_path)
        sheet = wb.sheets['ironandsteelmakertool']

        sheet.range('D16').value = request.POST.get('start_year')
        sheet.range('D17').value = request.POST.get('base_year_activity')
        sheet.range('D18').value = request.POST.get('base_year_emission')
        sheet.range('D20').value = request.POST.get('end_year')
        sheet.range('D21').value = request.POST.get('target_activity')
        sheet.range('D22').value = request.POST.get('target_year_output')
        sheet.range('D24').value = request.POST.get('scrap_base')
        sheet.range('D25').value = request.POST.get('scrap_target')
        # print("Start Year:", request.POST.get('start_year'))
        # print("Base Year Activity:", request.POST.get('base_year_activity'))
        # print("Base Year Emission:", request.POST.get('base_year_emission'))
        # print("End Year:", request.POST.get('end_year'))
        # print("Target Activity:", request.POST.get('target_activity'))
        # print("Target Year Output:", request.POST.get('target_year_output'))
        # print("Scrap Base:", request.POST.get('scrap_base'))
        # print("Scrap Target:", request.POST.get('scrap_target'))

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
        current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        image_name = f'{current_time}.png'
        server_image_path = os.path.join(save_directory, image_name)

        if screenshot is not None:
            screenshot.save(server_image_path)
            print(f"Content screenshot successfully saved to {server_image_path}")
        else:
            print("Failed to grab the image from clipboard.")

        wb.close()
        app.quit()
        image_url = os.path.join('img', image_name)  
        return render(request, 'steel.html', {'image_path': image_name})

    except Exception as e:
        if 'wb' in locals():
            wb.close()  # Ensure workbook is closed if an error occurs
        if 'app' in locals():
            app.quit()  # Ensure Excel app is closed if an error occurs
        return render(request, 'index.html', {'error': str(e)})

# formatting data
def format_dataframe(df):
    pd.set_option('display.float_format', lambda x: f'{x:.2f}')
    return df.to_html(index=False, na_rep='')

#displaying register form
def register(request):
    if request.method == 'POST':
        name = request.POST.get('name')
        firstname=request.POST.get('fname')
        lastname=request.POST.get('lname')
        phone = request.POST.get('phone')
        country = request.POST.get('country')
        email = request.POST.get('email')
        password = request.POST.get('password')

        if User.objects.filter(email=email).exists():
            messages.error(request, "Email is already registered")
        else:
            # Create user
            user = User.objects.create_user(username=name,first_name=firstname,last_name=lastname, email=email, password=password)
            user.save()

            # Create user profile with phone and country
            profile = Profile.objects.create(user=user, phone=phone, country=country)
            profile.save()

            messages.success(request, "Registration successful")
            # login(request, user)
            return redirect('login')

    return render(request, 'register.html')


#displaying login form
def user_login(request):
    if request.method == 'POST':
        email = request.POST.get('email')
        password = request.POST.get('password')

        # Fetch the user by email
        try:
            user = User.objects.get(email=email)
        except User.DoesNotExist:
            messages.error(request, "Invalid email or password.")
            return redirect('login')

        # Authenticate the user
        user = authenticate(request, username=user.username, password=password)
        if user is not None:
            login(request, user)
            return redirect('ste')  # Redirect to the steel page after login
        else:
            messages.error(request, "Invalid email or password.")

    return render(request, 'login.html')


# logout
def user_logout(request):
    auth_logout(request)  # Logout the user
    return redirect('login')  


# index contact form
def contact_view(request):
    if request.method == 'POST':
        form = ContactForm(request.POST)
        if form.is_valid():
            form.save()  # Save form data to the database
            messages.success(request, 'Your message has been sent successfully.')  # Use messages here
            return redirect('home')  # Redirect to the home page after success
        else:
            messages.error(request, 'There was an error with your submission. Please check the form.')
    
    # Optionally render the form even for GET requests
    form = ContactForm()  # Ensure form is instantiated for GET requests
    return render(request, 'index.html', {'form': form})