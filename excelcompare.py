from pyxlsb import open_workbook
from pathlib import Path

###PODAJ NAZWĘ FOLDERU Z TEST CASEM###
FOLDER_NAME = 'TestCase01'

### USTAWIENIA FORMATOWANIA KONSOLI ###
RED = "\033[1;4;31m"
GREEN = "\033[1;4;32m"
YELLOW = "\033[1;33m"
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
print('🔳 KluczZExcela - [lista wartości z excela] - [lista wartości z xml]\n')
## EXCEL ##

def konwert(el):
    excel_errors = {
        '0xf': f'❗BŁĄD SKŁADNI (#ARG!)❗',       # Kod dla #ARG! / #VALUE!
        '0x7': f'❗BŁĄD DZIELENIA (#DZIEL/0!)❗', # Kod dla #DIV/0!
        '0x17': f'❗BŁĄD BRAKU DANYCH (#N/D!)❗', # Kod dla #N/A
        '0x1c': f'❗BŁĄD NAZWY (#NAZWA?)❗'       # Kod dla #NAME?
    }

    if el in excel_errors:
        return excel_errors[el]

    try:
        number = round(float(el), 10)
        if number.is_integer():
            return int(number)
        return number
    except (ValueError, TypeError):
        return el

excel_dict = {}
excel_url = folder / NAME_EXCEL_FILE
with open_workbook(str(excel_url)) as wb:
    with wb.get_sheet('OUTBOUND') as sheet:
        for row in sheet.rows():
            if row[0].v is not None:
                klucz = row[0].v.split('.')[-1]
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
# for i in xml_list:
#     print(i)

### PORÓWNANIE PLIKÓW ###
md_content = []
founded_in_xml = ['Nazwa zmiennej']
for excel_key, excel_value in excel_dict.items():
    if excel_key == 'Nazwa zmiennej':
        continue
    open_mark = f"<{excel_key}>"
    close_mark = f"</{excel_key}>"
    xml_founded_value = []
    for xml in xml_list:
        if xml.startswith(open_mark) and xml.endswith(close_mark):
            xml_value = konwert(xml.split('>')[1].split('<')[0])
            xml_founded_value.append(xml_value)
            founded_in_xml.append(excel_key)

    exc = "  \|  ".join(str(i) for i in excel_value)
    xml = "  \|  ".join(str(i) for i in xml_founded_value)
    if not xml_founded_value:
        print(f'⚠️ {YELLOW}{excel_key} - {excel_value} - nie znaleziono w xml{RESET} ')

        md_content.append(f'| ⚠️ | {excel_key} | {exc} | nie znaleziono w xml |')
    elif excel_value == xml_founded_value:
        print(f'✅ {GREEN}{excel_key} - {excel_value} - {xml_founded_value}{RESET} ')

        md_content.append(f'| ✅ | {excel_key} | {exc} | {xml} |')
    else:
        print(f'❌ {RED}{excel_key} - {excel_value} - {xml_founded_value}{RESET} ')
        md_content.append(f'| ❌ |{excel_key} | {exc} | {xml} |')



print('\nBrakujące klucze w pliku xml:')
brakujace_klucze = set(excel_dict.keys()) - set(founded_in_xml)
print(brakujace_klucze)

### RAPORT ###
import datetime
md_content = [
    "## 📊 RAPORT TESTU REGRESYJNEGO API",
    f"+ Przypadek testowy: {FOLDER_NAME}",
    f"+ Data sporządzenia raportu: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
    "#### Porównanie wartości kluczy\n",
    f"| 🎭 | KLUCZ | PLIK EXCEL | PLIK XML |",
    "| :--- | :--- | :--- | :--- |",
] + md_content

with open(f'{FOLDER_NAME}/RAPORT.md', 'w', encoding='utf-8') as raport:
    raport.write("\n".join(md_content))