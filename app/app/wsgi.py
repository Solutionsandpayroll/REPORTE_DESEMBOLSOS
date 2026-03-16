import os
import sys
from django.core.wsgi import get_wsgi_application


path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if path not in sys.path:
    sys.path.insert(0, path)

# Ahora que la carpeta interna APP está en el path, 
# podemos usar 'app.settings' directamente.
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'app.settings')

application = get_wsgi_application()
app = application # Vercel a veces busca la variable 'app'