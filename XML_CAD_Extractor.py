import os
from lxml import etree
import json

def get_config():
    if "config.json" not in os.listdir():
        json_data = {
            "YES": "y",
            "NO": "n",
            "YES_NO": "yn",
            "WO_FILE": "wo_list.csv",
            "YN": "(y/n)",
            "XML_FILES_FOLDER": "XML_Files",
            "XML_DESTINATION_URL": "\\\\testurl.intranet.testuser.com\\FABRICACION\\DXFS",
            "XML_OLD_FOLDER": "XML_Old",
            "DEFAULT_OUTPUT_FILENAME": "output.xml",
            "DEFAULT_OUTPUT_FOLDER": "XML_Output"
        }
        
        with open("config.json", "w", encoding='utf-8') as config_file:
            json.dump(json_data, config_file, ensure_ascii=False, indent=4)
        return json_data
    
    else:
        with open("config.json") as config_file:
            return json.load(config_file)
        
config = get_config()

YES = config["YES"]
NO = config["NO"]
YES_NO = config["YES_NO"]
WO_FILE = config["WO_FILE"]
YN = config["YN"]
XML_FILES_FOLDER = config["XML_FILES_FOLDER"]
XML_OLD_FOLDER = config["XML_OLD_FOLDER"]
XML_DESTINATION_URL = config["XML_DESTINATION_URL"]
DEFAULT_OUTPUT_FILENAME = config["DEFAULT_OUTPUT_FILENAME"]
DEFAULT_OUTPUT_FOLDER = config["DEFAULT_OUTPUT_FOLDER"]

def write_orders_to_file():
    '''Funkce pro manuální zápis WO do listu a souboru'''
    orders = []
    while True:
        write_orders = input(f'Přejete si do souboru "wo_list" zapsat požadované zakázky? {YN} (prázdná hodnota přeskočí zápis).: ')
        if not write_orders or write_orders.lower() in YES_NO:
            break
        else:
            print("Neplatná hodnota.")
    if write_orders.lower() == YES:
        with open(WO_FILE, "w") as new_list:
            print('Vkládejte jednotlivé zakázky a každou z nich potvrďte klávesou "Enter", po poslední zakázce potvrďte prázdnou hodnotu pro dokončení:')
            while True:
                try:
                    order = input("")
                    if not order:
                        if not orders:
                            print("Nezadali jste žádné zakázky!")
                        else:
                            print(f"Do souboru bude zapsáno {len(orders)} zakázek")
                        break
                    orders.append(int(order))
                except ValueError:
                    print("Číslo zakázky musí být v číselném formátu!")
            for order in orders:
                new_list.write(f"{order}\n")
    return orders

def get_wo_list():
    """Funkce na vytažení WO ze souboru"""
    wo_list = []
    try:
        with open(WO_FILE, "r") as wo_file:
            for index, line in enumerate(wo_file):
                try:
                    wo_list.append(int(line.strip()))
                except ValueError:
                    if not line.strip():
                        continue
                    else:
                        print(f"Nesprávná hodnota na řádku č.{index + 1}.\nExtrakce pokračuje bez ní.")
        return wo_list
    except FileNotFoundError:
        while True:
            create_file = input(f'Soubor "{WO_FILE}" nenalezen, přejete si jej vytvořit? {YN} (prázdná hodnota tvorbu přeskočí a ukončí program).: ')
            if not create_file or create_file.lower() in YES_NO:
                break
            else:
                print("Neplatná hodnota.")
        if create_file.lower() == YES:
            with open(WO_FILE, "w") as new_list:
                pass
            return wo_list

def get_xml_list():
    '''Funkce na vytažení XML souborů ze složky specifikované konstantou XML_FILES_FOLDER'''
    xml_list = []
    if XML_FILES_FOLDER not in os.listdir():
        while True:
            make_dir = input(f'Složka {XML_FILES_FOLDER} nebyla nalezena, přejete si jí vytvořit? {YN} (prázdná hodnota také vytvoří složku).: ')
            if not make_dir or make_dir.lower() in YES_NO:
                break
            else:
                print("Neplatná hodnota.")
        if make_dir.lower() != NO:
            os.mkdir(XML_FILES_FOLDER)
            input(f'Složka {XML_FILES_FOLDER} byla vytvořena, pro správný běh programu do ní vložte XML soubory s požadovanými daty.\nPo vložení souborů stiskni "Enter" pro pokračování')
    if not [file for file in os.listdir(XML_FILES_FOLDER) if file.endswith(".xml")]:
        input(f'Ve složce {XML_FILES_FOLDER} chybí XML soubory! Vložte je zde nyní a stiskněte "Enter".')
    for line in os.listdir(XML_FILES_FOLDER):
        if not line.strip().endswith(".xml"):
            print(f'Soubor "{line.strip()}" není typu XML a proto nebude použit.')
            continue
        else:
            xml_list.append(line.strip())
    return xml_list

def extract_xml_data(wo_list, xml_list):
    '''Funkce pro vytažení dat z XML souborů'''
    xml_data = {}
    not_found = []
    for file in xml_list:
        xml_tree = etree.parse(f"{XML_FILES_FOLDER}/{file}")
        xml_root = xml_tree.getroot()
        for production_order_element in xml_root.xpath('//ProductionOrder'):
            wo_num = production_order_element.get('OrderNo')
            if int(wo_num) not in wo_list:
                continue
            else:
                part_no = production_order_element.xpath("./PartNo/text()")[0]
                material = production_order_element.xpath("./Material/text()")[0]
                xml_data[int(wo_num)] = [part_no, material]
    for wo in wo_list:
        if wo not in xml_data.keys():
            not_found.append(wo)
    return xml_data, not_found

def write_missing_wo_list(not_found):
    '''Interaktivní funkce co nabídne, jestli se má zapsat soubor s chybějícími WO'''
    if not not_found:
        print("Všechny WO byly nalezeny!")
    else:
        not_found_str = '\n'.join(list(map(str, sorted(not_found))))
        print(f"Tyto WO nebyly v XML souborech nalezeny:\n{not_found_str}\n")
        while True:
            write_file = input(f'Chcete zapsat chybějící WO do souboru "missing_list.csv" {YN} (prázdná hodnota také přeskočí zápis).: ')
            if not write_file or write_file.lower() in YES_NO:
                break
            else:
                print("Neplatná hodnota.")
        if write_file.lower() == YES:
            with open("missing_list.csv", "w") as missing_file:
                for wo in not_found:
                    missing_file.write(f"{str(wo)}\n")
            print('Soubor "missing_list.csv" úspěšně vytvořen.')

def write_new_import_file(xml_data):
    '''Funkce pro zápis dat do nového XML souboru'''
    def choose_filename():
        '''Vnořená funkce pro zjištění dostupnosti názvu'''
        if DEFAULT_OUTPUT_FOLDER not in os.listdir():
            os.mkdir(DEFAULT_OUTPUT_FOLDER)
        folder = DEFAULT_OUTPUT_FOLDER
        filename = DEFAULT_OUTPUT_FILENAME
        replace = NO
        while filename in os.listdir(DEFAULT_OUTPUT_FOLDER) and replace == NO:
            while True:
                overwrite = input(f'Soubor "{filename}" již existuje, chcete jej přepsat? {YN} (prázdná hodnota a "n" navrhne změnu jména) : ')
                if not overwrite or overwrite.lower() in YES_NO:
                    break
                else:
                    print("Neplatná hodnota.")
            if overwrite.lower() != YES:
                filename = input("Vložte nový název výstupního souboru (bez přípony): ") + ".xml"
            else:
                replace = YES
        return (folder, filename)

    new_root = etree.Element("PPSImport", Version="1.1")
    new_parts = etree.SubElement(new_root, "Parts")

    for data in xml_data.values():
        part_element = etree.SubElement(new_parts, "Part", PartNo=data[0])
        cad_filename = etree.SubElement(part_element, "CADFilename")
        cad_filename.text = f"\{XML_DESTINATION_URL}{data[0]}.dxf"
        material = etree.SubElement(part_element, "Material")
        material.text = data[1]
    
    xml_tree = etree.ElementTree(new_root)
    while True:
        folder, filename = choose_filename()
        try:
            xml_tree.write(f"{folder}/{filename}", encoding="utf-8", xml_declaration=True, pretty_print=True)
            break
        except OSError:
            print("Nevyhovující název souboru!")
    print(f'Vytvoření XML souboru "{filename}" proběhlo v pořádku')

def move_xml_files(xml_list):
    '''Funkce pro přesunutí XML souborů do složky specifikované konstantou XML_OLD_FOLDER po dokončení extrakce'''
    while True:
        move_files = input(f'Chcete použité XML soubory přesnuout do složky {XML_OLD_FOLDER}? {YN} (prázdná hodnota soubory přesune) : ')
        if not move_files or move_files.lower() in YES_NO:
            break
        else:
            print("Neplatná hodnota.")
    if move_files.lower() != NO:
        if XML_OLD_FOLDER not in os.listdir():
            print(f'Složka {XML_OLD_FOLDER} nebyla nalezena a bude vytvořena.')
            os.mkdir(XML_OLD_FOLDER)
        for file in xml_list:
            os.replace(f'{XML_FILES_FOLDER}\{file}',f'{XML_OLD_FOLDER}\{file}')
        print("všechny soubory byly přesunuty.")
    else:
        print("Soubory nebudou přesunuty.")

def main():
    try:

        wo_list = get_wo_list() #Pokus o extrakci WO
        if not wo_list and WO_FILE in os.listdir():
            print(f'Soubor "{WO_FILE}" je prázdný!')
            wo_list = write_orders_to_file()
            if not wo_list:
                raise FileNotFoundError
        elif WO_FILE not in os.listdir():
            print(f'Soubor "{WO_FILE}" nebyl vytvořen!')
            raise FileNotFoundError

        wo_list_str = "\n".join(list(map(str, sorted(wo_list))))
        print(f"\nnalezené WO:\n{wo_list_str}\n")

        xml_list = get_xml_list() #Pokus o extrakci XML
        if not xml_list:
            print(f'Nevložili jste žádné XML soubory do složky "{XML_FILES_FOLDER}".')
            raise FileNotFoundError
        xml_list_str = "\n".join(xml_list)
        print(f"nalezené XML:\n{xml_list_str}\n")
        xml_data, not_found = extract_xml_data(wo_list, xml_list) #Pokus o extrakci dat z XML
        if not xml_data:
            print("XML soubory neobsahují správná data!")
            raise FileNotFoundError
        
        write_missing_wo_list(not_found)
        write_new_import_file(xml_data)
        move_xml_files(xml_list)

        input(f'Extrakce dokončena, stiskni "Enter" pro ukončení.')

    except FileNotFoundError:
        input('Program selhal z důvodu chybějících dat, stiskni "Enter" pro ukončení.')

if __name__ == "__main__":
    main()