import os
from django.core.wsgi import get_wsgi_application

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'app.settings') # Verifica que aquí diga 'app.settings'

application = get_wsgi_application()
app = application  # <--- ESTA LÍNEA ES CLAVE PARA VERCEL