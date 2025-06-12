from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('get_report/', views.get_report, name='get_report'),
]
