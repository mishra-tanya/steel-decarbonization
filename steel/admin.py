from django.contrib import admin

# Register your models here.
from .models import ContactMessages  
from .models import Profile

admin.site.register(ContactMessages)
admin.site.register(Profile)