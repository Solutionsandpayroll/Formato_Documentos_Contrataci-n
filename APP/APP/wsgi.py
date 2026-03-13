import os
import sys
from django.core.wsgi import get_wsgi_application

# Buscamos la ruta de la primera carpeta 'APP' (donde está manage.py)
# __file__ es .../APP/APP/wsgi.py
# parent es .../APP/APP/
# grandparent es .../APP/ (Esta es la que queremos en el path)
path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if path not in sys.path:
    sys.path.insert(0, path)

# Ahora que la carpeta interna APP está en el path, 
# podemos usar 'APP.settings' directamente.
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'APP.settings')

application = get_wsgi_application()
app = application # Vercel a veces busca la variable 'app'