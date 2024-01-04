* El archivo zip tiene que estar en el mismo directorio que el script de Python
* No cambiar el nombre del archivo zip, debe ser: victims.zip
* El resultado se muestra como test.cv
* Ejecutar los siguientes pasos en una Powershell, desde el mismo directorio donde se encuentra el script.
``` 
python -m venv .venv
.\.venv\Scripts\activate
pip install -r req.txt
python test.py
``` 