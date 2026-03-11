from django.urls import path
from . import views

urlpatterns = [
    path('', views.subir_excel, name='subir_excel_view'), # Primera pantalla
    path('colaboradores/', views.generar_word, name='generar_word'), # Segunda pantalla
]