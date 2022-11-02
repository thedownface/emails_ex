from django.urls import path

from . import views

app_name = "e_selenium"

urlpatterns = [
    path('e_selenium/', views.e_selenium, name='e_selenium'),
]