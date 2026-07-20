from pyxlsb import open_workbook
from pathlib import Path
import json

### USTAWIENIA FORMATOWANIA KONSOLI ###
RED = "\033[1;4;31m"
GREEN = "\033[1;4;32m"
UNDERLINE = "\033[4m"
RESET = "\033[0m"

### POBIERANIE PLIKÓW Z KATALOGÓW ###
print(f'{UNDERLINE}WELCOME TO EXCEL-FILES COMPARER{RESET}\nPreparing a files...')

try:
    data = json.loads(Path('DATA.json').read_text(encoding='utf-8'))
    FOLDER_NAME = data['FOLDER_NAME']
except(FileNotFoundError, KeyError):
    print(f'{RED}Problem with file DATA.json{RESET}')
    exit()

folder = Path(FOLDER_NAME)

if folder.exists():
    print (f'{GREEN}LOADED FOLDER CORRECTLY{RESET}')
    print(folder)
else:
    print (f'{RED}FOLDER {folder} DOESN\'T EXIST{RESET}')
    exit()

excel_file = [i.name for i in folder.glob('*.xlsb')]
if len(excel_file) == 1:
    print(f'{GREEN}LOADED 1 EXCEL FILE CORRECTLY{RESET}')
    NAME_EXCEL_FILE = excel_file[0]
    print(NAME_EXCEL_FILE)
else:
    print(f'{RED}ERROR WITH LOADING EXCEL FILE\nPlease, check that in folder is only 1 file.{RESET}')
    exit()

xml_file = [i.name for i in folder.glob('*.xml')]
if len(xml_file) == 1:
    print(f'{GREEN}LOADED 1 XML FILE CORRECTLY{RESET}')
    NAME_XML_FILE = xml_file[0]
    print(NAME_XML_FILE)
else:
    print(f'{RED}ERROR WITH LOADING XML FILE\nPlease, check that in folder is only 1 file.{RESET}')
    exit()

### ODCZYTYWANIE DANYCH ###
## EXCEL ##
excel_dict = {}
excel_url = folder / NAME_EXCEL_FILE
with open_workbook(str(excel_url)) as wb:
    with wb.get_sheet('OUTBOUND') as sheet:
        for row in sheet.rows():
            if row[0].v is not None:
                excel_dict[(row[0].v).removeprefix('outbound.')] = row[1].v

for i in excel_dict.items():
    print(i)

## XML ##
xml_url = folder / NAME_XML_FILE
with open(xml_url, 'r', encoding='utf-8') as xml_file:
    xml_txt = xml_file.read()

xml_txt = xml_txt.split('<m:xmlString>&lt;outbound>')[1].split('</m:xmlString>')[0].strip()
xml_list = xml_txt.split('\n')
xml_list = [i.strip().replace('&lt;','<') for i in xml_list]
for i in xml_list:
    print(i)