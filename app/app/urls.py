from django.contrib import admin
from django.urls import path
from generador_plano import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('',           views.index,           name='index'),
    path('preview/',   views.preview_datos,   name='preview'),
    path('generar/',   views.generar_archivo,  name='generar'),
]
