# parse_dgii_casos_prueba
Herramienta para generar inserts de sql de acuerdo a casos de prueba de la dgii en el proceso de verificacion y a la estructura de base de datos de la documentacion oficial

## how to use

1. Instalar la ultima version de python compatible con openpyxl actualmente probada contra Python 3.13.0
2. create a virtual environment
3. install the requirements from the requirements file in requirements.txt
4. run the script parse_dgii.py

```bash
python -m venv .venv
. ./.venv/bin/activate
# or on windows
# .\.venv\Scripts\activate
pip install --upgrade pip
pip install -r requirements.txt

python parse_dgii.py __name_of_file_excel_or_csv__
```

El archivo puede ser un xlsx o un csv separado por '|'. y el resultado se guardara en un archivo res.sql
