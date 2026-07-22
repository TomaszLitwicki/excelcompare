from pyxlsb import open_workbook
from pathlib import Path

### STWÓRZ FOLDER 'xml' i tam wgraj pliki xml ###
### PLIK EXCELA UMIEŚĆ W TYM SAMYM KATALOGU CO PLIK PYTHONA ###
### USTAWIAJ W EXCELU INTERESUJĄCY CIĘ PRZYPADEK TESTOWY ###
### PODAWAJ TUTAJ NAZWĘ PLIKU XML DO ANALIZY###
XML_FILE_NAME = 'nazwa_pliku_xml'



### RESZTA DZIAŁA JUŻ SAMA ###

### USTAWIENIA FORMATOWANIA KONSOLI ###
RED = "\033[1;4;31m"
GREEN = "\033[1;4;32m"
YELLOW = "\033[1;33m"
UNDERLINE = "\033[4m"
RESET = "\033[0m"

### POBIERANIE PLIKÓW Z KATALOGÓW ###
print(f'{UNDERLINE}WELCOME TO EXCEL-FILES COMPARER{RESET}\nPreparing a files...')

BASE_DIR = Path(__file__).resolve().parent
folder_xml = BASE_DIR / 'xml'
folder_reports = BASE_DIR / 'reports'


if folder_xml.exists():
    print (f'{GREEN}LOADED FOLDER CORRECTLY{RESET}')
    print(folder_xml)
else:
    print (f'{RED}FOLDER {folder_xml} DOESN\'T EXIST. {RESET}\n{YELLOW}Please, create a folder xml and put there a files xml{RESET}')
    exit()

excel_file = [i.name for i in BASE_DIR.glob('*.xlsb') if not i.name.startswith('~$')]
if len(excel_file) == 1:
    print(f'{GREEN}LOADED 1 EXCEL FILE CORRECTLY{RESET}')
    NAME_EXCEL_FILE = excel_file[0]
    print(NAME_EXCEL_FILE)
else:
    print(f'{RED}ERROR WITH LOADING EXCEL FILE\nPlease, check that in folder where is the Python file is only 1 excel file too.{RESET}')
    exit()

if not XML_FILE_NAME.endswith('.xml'):
    XML_FILE_NAME = f'{XML_FILE_NAME}.xml'
xml_files = [i.name for i in folder_xml.glob('*.xml')]
if XML_FILE_NAME in xml_files:
    print(f'{GREEN}LOADED 1 XML FILE CORRECTLY{RESET}')
    print(XML_FILE_NAME)
else:
    print(f'{RED}ERROR WITH LOADING XML FILE{RESET}\n{YELLOW}Please, check that in folder xml is file ({XML_FILE_NAME}).{RESET}')
    exit()

### ODCZYTYWANIE DANYCH ###
print('\nAnaliza danych...\n')
print('🔳 KluczZExcela - [lista wartości z excela] - [lista wartości z xml]')

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
excel_url = folder_xml / NAME_EXCEL_FILE
sections = set()
with open_workbook(str(NAME_EXCEL_FILE)) as wb:
    with wb.get_sheet('OUTBOUND') as sheet:
        for row in sheet.rows():
            cell_a_val = row[0].v
            cell_b_val = row[1].v
            if cell_a_val is not None:
                klucz: str = row[0].v
                klucz_lis = (row[0].v.split('.'))
                if len(klucz_lis) > 1:
                    sections.add(klucz_lis[-2])
                    outbound = klucz_lis[0]
                if cell_b_val is not None:
                    wartosc = str(row[1].v).split('|')
                else:
                    wartosc = '🚫'
                przekonwert = [konwert(i) for i in wartosc]
                excel_dict[klucz] = przekonwert

## XML ##
xml_url = folder_xml / XML_FILE_NAME
with open(xml_url, 'r', encoding='utf-8') as xml_file:
    xml_txt = xml_file.read()

xml_txt = xml_txt.split('<m:xmlString>&lt;outbound>')[1].split('</m:xmlString>')[0].strip().replace('&lt;','<')
xml_outbound = xml_txt
xml_secionizer = {}

for sec in sections.copy():
    if f'<{sec}/>' in xml_txt:
        sections.remove(sec)
        xml_secionizer[sec] = ['']

for sec in sections:
    if sec != outbound:
        xml_secionizer[sec] = [i.strip() for i in xml_txt.split(f'<{sec}>')[1].split(f'</sec>')[0].strip().split('\n')]
    else:
        for s in sections:
            if s != outbound:
                l, c = xml_outbound.split(f'<{s}>')
                c, p = c.strip().split(f'</{s}>')
                xml_outbound = l + p.strip()
        xml_secionizer[sec] = [i.strip() for i in xml_outbound.split('\n')]

xml_list = xml_txt.split('\n')
xml_list = [i.strip() for i in xml_list]

### PORÓWNANIE PLIKÓW ###
md_content = []
founded_in_xml = ['Nazwa zmiennej']
for excel_key, excel_value in excel_dict.items():
    if excel_key == 'Nazwa zmiennej':
        continue
    excel_key_list = excel_key.split('.')
    excel_section = excel_key_list[-2]
    open_mark = f"<{excel_key_list[-1]}>"
    close_mark = f"</{excel_key_list[-1]}>"
    xml_founded_value = []
    xml_list: list = xml_secionizer[excel_section]
    for xml in xml_list:
        if xml.startswith(open_mark) and xml.endswith(close_mark):
            xml_value = konwert(xml.split('>')[1].split('<')[0])
            xml_founded_value.append(xml_value)
            founded_in_xml.append(excel_key)

    exc = "  &#124;  ".join(str(i) for i in excel_value)
    xml = "  &#124;  ".join(str(i) for i in xml_founded_value)
    if not xml_founded_value:
        print(f'⚠️ {YELLOW}{excel_key} - {excel_value} - nie znaleziono w xml{RESET} ')

        md_content.append(f'| ⚠️ | {excel_key.split(".", maxsplit=1)[1]} | {exc} | nie znaleziono w xml |')
    elif excel_value == xml_founded_value:
        print(f'✅ {GREEN}{excel_key} - {excel_value} - {xml_founded_value}{RESET} ')

        md_content.append(f'| ✅ | {excel_key.split(".", maxsplit=1)[1]} | {exc} | {xml} |')
    else:
        print(f'❌ {RED}{excel_key} - {excel_value} - {xml_founded_value}{RESET} ')
        md_content.append(f'| ❌ |{excel_key.split(".", maxsplit=1)[1]} | {exc} | {xml} |')



print('\nBrakujące klucze w pliku xml:')
brakujace_klucze = set(excel_dict.keys()) - set(founded_in_xml)
print(brakujace_klucze)

### RAPORT ###
import datetime
md_content = [
    "## 📊 RAPORT TESTU REGRESYJNEGO API",
    f"+ Przypadek testowy: {XML_FILE_NAME.removesuffix('.xml')}",
    f"+ Data sporządzenia raportu: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
    "#### Porównanie wartości kluczy\n",
    f"| 🎭 | KLUCZ | PLIK EXCEL | PLIK XML |",
    "| :--- | :--- | :--- | :--- |",
] + md_content

folder_reports.mkdir(parents=True, exist_ok=True)
with open(f'{folder_reports}/{XML_FILE_NAME.removesuffix(".xml")}.md', 'w', encoding='utf-8') as raport:
    raport.write("\n".join(md_content))