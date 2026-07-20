from pyxlsb import open_workbook
from pathlib import Path

###PODAJ NAZWĘ FOLDERU Z TEST CASEM###
FOLDER_NAME = 'TestCase01'

### USTAWIENIA FORMATOWANIA KONSOLI ###
RED = "\033[1;4;31m"
GREEN = "\033[1;4;32m"
UNDERLINE = "\033[4m"
RESET = "\033[0m"

### POBIERANIE PLIKÓW Z KATALOGÓW ###
print(f'{UNDERLINE}WELCOME TO EXCEL-FILES COMPARER{RESET}\nPreparing a files...')

folder = Path(FOLDER_NAME)

if folder.exists():
    print (f'{GREEN}LOADED FOLDER CORRECTLY{RESET}')
    print(folder)
else:
    print (f'{RED}FOLDER {folder} DOESN\'T EXIST{RESET}')
    exit()

excel_file = [i.name for i in folder.glob('*.xlsb') if not i.name.startswith('~$')]
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
print('\nAnaliza danych...')
print('Po lewej stronie jest klucz z excela - [w nawiasach jest pożądana wartość z excela] - po prawej stronie jest wartość z xmla\n')
## EXCEL ##

def konwert(el):
    try:
        liczba = float(el)
        if liczba.is_integer():
            return int(liczba)
        return liczba
    except ValueError:
        return el

excel_dict = {}
excel_url = folder / NAME_EXCEL_FILE
with open_workbook(str(excel_url)) as wb:
    with wb.get_sheet('OUTBOUND') as sheet:
        for row in sheet.rows():
            if row[0].v is not None:
                klucz = (row[0].v).split('.')[-1]
                wartosc = str(row[1].v).split('|')
                przekonwert = [konwert(i) for i in wartosc]
                excel_dict[klucz] = przekonwert

# for i in excel_dict.items():
#     print(i)

## XML ##
xml_url = folder / NAME_XML_FILE
with open(xml_url, 'r', encoding='utf-8') as xml_file:
    xml_txt = xml_file.read()

xml_txt = xml_txt.split('<m:xmlString>&lt;outbound>')[1].split('</m:xmlString>')[0].strip()
xml_list = xml_txt.split('\n')
xml_list = [i.strip().replace('&lt;','<') for i in xml_list]
for i in xml_list:
    print(i)

### PORÓWNANIE PLIKÓW ###
founded_in_xml = ['Nazwa zmiennej']
for excel in excel_dict.items():
    excel_klucz = excel[0]
    excel_wartosc = excel[1]
    for xml in xml_list:
        if f'<{excel_klucz}>' in xml or f'<{excel_klucz}/>' in xml:
            xml_wartosc = konwert(xml.split('>')[1].split('<')[0])
            if xml_wartosc in excel_wartosc:
                print(f'{GREEN}{excel_klucz} - {excel_wartosc} - {xml_wartosc}{RESET}')
            else:
                print(f'{RED}{excel_klucz} - {excel_wartosc} - {xml_wartosc}{RESET}')

            founded_in_xml.append(excel_klucz)

print('\nBrakujące klucze w pliku xml:')
brakujace_klucze = set(excel_dict.keys()) - set(founded_in_xml)
print(brakujace_klucze)