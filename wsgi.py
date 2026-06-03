import sys
import os

# Caminho do projeto no PythonAnywhere (ajuste SEU_USUARIO)
path = os.path.dirname(os.path.abspath(__file__))
if path not in sys.path:
    sys.path.insert(0, path)

from a2wsgi import ASGIMiddleware
from main import app

application = ASGIMiddleware(app)
