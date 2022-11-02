from django.urls import path

from . import views

app_name = "e_scrapy"

urlpatterns = [
  
    path('scrape/',views.e_scrapy, name='e_scrapy')
]
