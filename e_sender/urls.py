from django.urls import path

from . import views

app_name = "e_sender"

urlpatterns = [
    path('e_sender/', views.e_sender, name='e_sender')
]