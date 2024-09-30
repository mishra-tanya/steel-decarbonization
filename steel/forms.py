# forms.py
from django import forms
from .models import ContactMessages

class FilterForm(forms.Form):
    start_date = forms.DateField(required=False, widget=forms.TextInput(attrs={'type': 'date'}))
    end_date = forms.DateField(required=False, widget=forms.TextInput(attrs={'type': 'date'}))
    category = forms.CharField(max_length=100, required=False)


class ContactForm(forms.ModelForm):
    class Meta:
        model = ContactMessages
        fields = ['name', 'contact', 'email', 'subject', 'message']

