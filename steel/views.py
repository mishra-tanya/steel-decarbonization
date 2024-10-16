from django.shortcuts import render, redirect
import pandas as pd
import os
from django.conf import settings
from django.conf import settings
from django.shortcuts import render
from openpyxl import load_workbook
import xlwings as xw
import openpyxl
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
        numeric_data_1 = sheet.range('B53:H56').value
        df1 = pd.DataFrame(numeric_data_1)
        df1 = df1.dropna(how='all')
        df1 = df1.dropna(axis=1, how='all') 
        df1 = df1.fillna('')  

        numeric_data_2 = sheet.range('F94:AD99').value
        df2 = pd.DataFrame(numeric_data_2)

       
        df2 = df2.dropna(how='all')  
        df2 = df2.dropna(axis=1, how='all') 
        df2 = df2.fillna('')
        
        table_html_1 = format_dataframe(df1)
        table_html_2 = format_dataframe(df2)

        save_directory = os.path.join(settings.BASE_DIR, 'static/img')
        os.makedirs(save_directory, exist_ok=True) 

        image_paths = []

       
        first_range = sheet.range('A32:H50')
        first_range.copy()
        time.sleep(1)
        screenshot1 = ImageGrab.grabclipboard()
        if screenshot1:
            current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            image_name1 = f'{current_time}_1.png'
            server_image_path1 = os.path.join(save_directory, image_name1)
            screenshot1.save(server_image_path1)
            image_paths.append(image_name1)

       
        second_range = sheet.range('A58:L84')
        second_range.copy()
        time.sleep(1)
        screenshot2 = ImageGrab.grabclipboard()
        if screenshot2:
            current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            image_name2 = f'{current_time}_2.png'
            server_image_path2 = os.path.join(save_directory, image_name2)
            screenshot2.save(server_image_path2)
            image_paths.append(image_name2)

   
        numeric_data = sheet.range('A94:AD99').value

       
        wb.close()
        app.quit()

        # Render the template with image paths and numeric data
        return render(request, 'steel.html', {
            'image1': image_paths[0] if len(image_paths) > 0 else None,
            'image2': image_paths[1] if len(image_paths) > 1 else None,
            'table_html_1': table_html_1,
            'table_html_2': table_html_2,
        })

    except Exception as e:
        if 'wb' in locals():
            wb.close()  # Ensure workbook is closed if an error occurs
        if 'app' in locals():
            app.quit()  # Ensure Excel app is closed if an error occurs
        return render(request, 'index.html', {'error': str(e)})
    

# formatting data
def format_dataframe(df):

    formatted_df = df.copy()

    formatted_df.iloc[1:] = formatted_df.iloc[1:].applymap(
        lambda x: f'{x:.2f}' if isinstance(x, float) else (f'{int(x)}' if isinstance(x, int) else x)
    )

    return formatted_df.to_html(index=False, header=False, na_rep='', classes='styled-table')

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