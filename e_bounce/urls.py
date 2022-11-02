from django.urls import path

from . import views

app_name = "e_bounce"

urlpatterns = [
  
    path('e_bounce/',views.e_bounce, name='e_bounce')
]
