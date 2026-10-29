import sys
from pathlib import Path

# Permite "import generar_datos" desde tests/.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
