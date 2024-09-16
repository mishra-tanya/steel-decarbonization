from django.urls import path
from steel import views
from . import views

urlpatterns = [
    path('', views.index, name='home'),
    path('steel-decarbonisation/', views.steel, name='steel'),
    path('term-conditions/', views.term, name='terms'),

]