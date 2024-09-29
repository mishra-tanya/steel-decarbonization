from django.urls import path
from steel import views
from . import views

urlpatterns = [
    path('', views.index, name='home'),
    path('steel-decarbonisation/', views.steel, name='steel'),
    path('steel/', views.ste, name='ste'),
    path('term-conditions/', views.term, name='terms'),
    path('register/', views.register, name='register'),
    path('login/', views.login, name='login'),
    path('contact/', views.contact_view, name='contact'),

]