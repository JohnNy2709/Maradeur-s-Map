import ctypes
import os
import openpyxl
import datetime
import pandas
import nmap

# Сделать единообразными таблицы вланов
# Требования: ip и mac столбцы всегда называются IP и MAC соответственно
# Ip и mac столбцы всегда находятся на одной и той же позиции относительно других столбцов (например, всегда 1 и 2 столбцы)
# во всех листах таблицы должны быть полными. Не допускается пропуск ip адресов

kernel32 = ctypes.windll.kernel32
kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)

def print_hello_words():
    text = "\n\tMessrs Moony, Wormtail, Padfoot, and Prongs \n\tPurveyors of Aids to Magical Mischief-Makers \n\tare proud to present\n\tTHE MARAUDER'S MAP\n\n"
    print('\033[33m {}'.format(text))

def clear_screen():
    os.system('cls')

def if_format_needed(file_name, format_needed):
    index_found = file_name.find(format_needed)
    if index_found == (-1):
        return False
    else:
        return True

def press_any():
    print('\n\t\033[33mPress any key to keep doing things..\033[0m \n')
    input()
    clear_screen()

def error_message(*code_index):
    match code_index:
        case '0':
            print('VLAN does not exist in .xslx file')
        case _:
            print('\n\t\033[33m..you are doing something wrong.\033[0m ')
    input()
    clear_screen()

def user_choosing(max_digit_possible):
    #принять на ввод цифру/число и проверить, что оно допустимо
    digit_from_input = str(input('\n\t\033[0m'))
    if (digit_from_input.isdigit()):
        if (int(digit_from_input) == (max_digit_possible + 1)):
            #refresh
            clear_screen()
            return 0
        if (int(digit_from_input) <= max_digit_possible and int(digit_from_input) > 0):
            return digit_from_input
        else:
            error_message()
            return 0
    else:
        error_message()
        return 0

def choose_xlsx_file():
    file_indexes = len(os.listdir())
    def display_files():
        print('\t\033[33mHere are the files from my directory.\n\tChoose a .xlsx file that is older than the Gods:\n')
        files = os.listdir()
        file_index = 1
        for file_name in files:
            current_string = '\033[0m \t' + str(file_index) + '\033[33m  ' + file_name 
            print(current_string)
            file_index += 1
        print('\t' + str(file_index) + '\033[33m  Refresh\033[33m')
        return files
    all_files = display_files()
    file_digit = user_choosing(file_indexes)
    if (file_digit == 0):
        choose_xlsx_file()
        return 0
    else:
        file_name = all_files[int(file_digit)-1]
        if(if_format_needed(file_name, '.xlsx')):
            print('\n\t\033[33mThe choice is made. I appreciate it.\033[0m\n')
            xlsx_file_path = ''
            xlsx_file_path = os.getcwd() +'\\' + all_files[int(file_digit)-1]
            return xlsx_file_path
        else:
            error_message()
            choose_xlsx_file()
            return 0

def choose_vlan():
    print('\n\t\033[33mWho are you, warrior?\n\tChoose the VLAN you want to check:\n\033[0m')
    arr_VLAN = ['Vlan21NO','Vlan100', 'Vlan108', 'Vlan110', 'Vlan133', 'Vlan192']
    vlan_id = 1
    for vlan in arr_VLAN:
        current_string = '\033[0m\t' + str(vlan_id) + '\033[33m ' + arr_VLAN[vlan_id-1] 
        print(current_string)
        vlan_id+=1
    chosen_vlan_id = int(user_choosing(len(arr_VLAN)))
    if chosen_vlan_id == 0:
        choose_vlan()
        return 0
    else:
        #возвращает название влана
        return arr_VLAN[chosen_vlan_id-1]

def get_name_for_files():
    now = datetime.datetime.now()
    current_time = now.strftime("%d-%m-%Y %H-%M")
    new_xlsx_filename = "MTUCI Switches " + current_time + ".xlsx"
    change_log_filename = "changelog" + current_time + ".txt"
    return new_xlsx_filename, change_log_filename

def macAddCheckUtility(nmapDictionary, xlsx_file_path, sheetname_vlanname, new_changelogfile_filename):
    #creates a changelog file & returns correct list of mac-add
    #def createDictFromNmap(nmap_file_path):
    #    def MacAddSelection(line):
    #        macAdd = line[13:30]
    #        return (macAdd)
    #    def IPSelection(line):
    #        for i, c in enumerate(line):
    #            if c.isdigit():
    #                ipStartingIndex = i
    #                break
    #        ipAdd = line[ipStartingIndex:len(line)-1]
    #        return (ipAdd)
    #    dictIpMac = {}
    #    fr = open(nmap_file_path, 'r')
    #    for line in enumerate(fr):
    #        if ("Nmap done" in line[1]):
    #            break
    #        if ("Nmap scan report" in line[1]):
    #            ipAdd = IPSelection(line[1])
    #        if ("MAC Address:" in line[1]):
    #            macAdd = MacAddSelection(line[1])
    #            dictIpMac[ipAdd] = macAdd
    #    fr.close()
    #    return (dictIpMac)   
    def createDictFromGoogleTable(xlsx_file_path, sheetname):
        #IP и MAC должны быть названиями столбцов. Название - первая ячейка в столбце
        dictIpMacExcess = {}
        dictIpMac = {}
        excel_data_df = pandas.read_excel(xlsx_file_path, sheet_name=sheetname, usecols=['IP', 'MAC'])
        dictIpMacExcess = excel_data_df.to_dict(orient='split')
        for i in range(0, len(dictIpMacExcess['data'])):
            dictIpMac[dictIpMacExcess['data'][i][0]] = dictIpMacExcess['data'][i][1]
        return dictIpMac
    def сompare_Dicts(nmapDictionary, googleDict):
        wrongMac, correctMac, missingMac = {}, {}, {}
        for hostIP in (nmapDictionary.keys()):
            nMapMac = nmapDictionary[hostIP]
            macAddInGoogle = googleDict.get(hostIP, 0)
            if macAddInGoogle == 0 or str(macAddInGoogle) == 'nan':
                missingMac[hostIP] = nMapMac
            else:
                if macAddInGoogle == nMapMac:
                    correctMac[hostIP] = nMapMac
                else:
                    wrongMac[hostIP] = [nMapMac, macAddInGoogle]
        return wrongMac, correctMac, missingMac
    def createChangelogFile(wrongMac, correctMac, missingMac, name):
        with open(name, 'w') as fw:
            fw.write('Google Sheet contains wrong MAC-Addresses for the following IP-Addresses (%d):\n\nIP\t\tCorrect MAC Address\t\told MAC\n'%len(wrongMac))
            for line in wrongMac:
                fw.write(str(line) + "\t" + str(wrongMac[line][0]) + "\t" + str(wrongMac[line][1]) + "\n")
            fw.write('\nGoogle Sheet does not contain MAC-Addresses for the following IP-Addresses (%d): \n\nIP\t\tMAC Address\n'%len(missingMac))
            for line in missingMac:
                fw.write(str(line) + "\t" + str(missingMac[line]) + '\n')
            fw.write('\nGoogle Sheet contains correct information for the following IP-Addresses (%d):\n\nIP\t\tMAC Address\n'%len(correctMac))
            for line in correctMac:
                fw.write(str(line) + "\t" + str(correctMac[line]) + '\n')
    def update_existing_return_google_and_dict_missing(googleTable, missingMac):
        _not_existing_mac_dict = {}
        currentIp = googleTable.keys()
        for ip in missingMac:
            if (ip in currentIp):
                googleTable[ip] = missingMac[ip]
            else:
                _not_existing_mac_dict[ip] = missingMac[ip]
        return googleTable, _not_existing_mac_dict
    #nMapDictionary = createDictFromNmap(nmap_file_path)
    googleTableDictionary = createDictFromGoogleTable(xlsx_file_path, sheetname_vlanname)
    wrongMacDict, correctMacDict, missingMacDict = сompare_Dicts(nmapDictionary, googleTableDictionary)
    googleTableDictionary, missingIpDict = update_existing_return_google_and_dict_missing(googleTableDictionary, missingMacDict)
    createChangelogFile(wrongMacDict, correctMacDict, missingMacDict, new_changelogfile_filename)
    #googleTableDictionary.update(missingMacDict)#обновляем исходный словарь mac, которых в нем не было
    for ip in wrongMacDict:
        googleTableDictionary[ip] = wrongMacDict[ip][0]
    return list(googleTableDictionary.values()), missingIpDict

def changes_writer(xlsx_file_path, sheet_name, list_to_write, new_xlsx_filename):
    _read_xlsx_file = openpyxl.load_workbook(xlsx_file_path)
    xlsx_sheet = _read_xlsx_file[sheet_name]
    index_of_mac_in_list = 0
    for row in xlsx_sheet.iter_rows(min_row = 2, min_col=2, max_col=2):
        #проходимся по всем его cell. 2 - номер столбца
        for cell in row:
            try: cell.value = list_to_write[index_of_mac_in_list]
            except: break
        index_of_mac_in_list+=1
    _read_xlsx_file.save(filename=new_xlsx_filename)

def ending(missingIp):
    print('\n\t\033[33mSuccess\033[0m')
    match len(missingIpDictionary):
        case 0:
            press_any()
        case _:
            print('\n\t\033[33m The following list contains information about those IP-addresses that are not in Google Sheet,\n\tbut the scanner detected devices behind IP. See "Missing Mac" in changelog\n')
            print(missingIpDictionary)

def scan_network_return_dict(network):
    #Здесь же можно вытянуть hostname(реализовано)
    # Нужно придумать, как его применить
    nm = nmap.PortScanner()
    nmapDict = {}
    ip_address_list = []
    nets_dict = {
    'Vlan21NO': '192.168.21.0/24', 'Vlan6': '10.0.6.0/24',
    'Vlan108':'10.0.8.0/24', 'Vlan100':'10.0.100.0/24',
    'Vlan110':'10.0.0.0/22', 'Vlan192':'192.168.0.0/24', 'Vlan133':'10.0.33.0/24'
    }   
    nmap_scan_dict = nm.scan(hosts=nets_dict[network], arguments='-n -sP -PE -PA21,23,80,3389')
    ip_without_mac = []
    for ip in nmap_scan_dict['scan'].keys():
        ip_address_list.append(ip)
        try:    macAdd = nmap_scan_dict['scan'][ip]['addresses']['mac']
        except:
            print('KeyError, scanned ip has no mac from nmap. The problem IP:', ip)
            ip_without_mac.append(ip)
            continue
        hostname = nmap_scan_dict['scan'][ip]['hostnames'][0]['name']
        nmapDict[ip] = macAdd
    return nmapDict, ip_without_mac


print_hello_words()
press_any()
xlsx_path = choose_xlsx_file()
vlanname = choose_vlan()  
nmapDict, ip_without_mac = scan_network_return_dict(vlanname)
edited_xlsx_filename, changelog_filename = get_name_for_files()
correct_mac_list, missingIpDictionary= macAddCheckUtility(nmapDict, xlsx_path, vlanname, changelog_filename)
changes_writer(xlsx_path, vlanname, correct_mac_list, edited_xlsx_filename)
ending(missingIpDictionary)
press_any()