Cómo ejecutarlo en Windows

Coloca los tres archivos en una carpeta de trabajo junto con inventario_ferreteria_garza.xlsx. Después, abre una terminal en esa carpeta.

Paso 1. Crear el entorno virtual

python -m venv .venv
.\.venv\Scripts\Activate.ps1

Paso 2. Instalar dependencias

pip install -r requirements.txt

Paso 3. Ejecutar el proceso

python resurtido.py

El programa generará los reportes de pedidos y excepciones, los borradores por proveedor y el archivo de registro de ejecución.