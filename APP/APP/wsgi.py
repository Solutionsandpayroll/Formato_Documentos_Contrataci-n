# En APP/APP/wsgi.py
import os
import sys
from django.core.wsgi import get_wsgi_application

# Esto sube dos niveles (hasta la raíz del proyecto) para que Python vea la primera carpeta /APP/
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'APP.APP.settings') # <--- DEBE SER APP.APP

application = get_wsgi_application()