# urls.py
from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('download/', views.download_file, name='download_file'),
    path('download-container/', views.download_container_folder, name='download_container_folder'),
]