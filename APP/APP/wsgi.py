import os
import sys
from django.core.wsgi import get_wsgi_application

# Esto agrega la raíz del proyecto al path de búsqueda de Python
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'APP.APP.settings')

application = get_wsgi_application()