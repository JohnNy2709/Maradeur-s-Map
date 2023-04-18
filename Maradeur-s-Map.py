import ctypes
import os
import openpyxl
import datetime
import pandas
import nmap
import pickle
from email.policy import default
kernel32 = ctypes.windll.kernel32
kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)

def print_hello_words():
    text = "\n\tMessrs Moony, Wormtail, Padfoot, and Prongs \n\tPurveyors of Aids to Magical Mischief-Makers \n\tare proud to present\n\tTHE MARAUDER'S MAP\n\n"
    print('\033[33m {}'.format(text))

def choose_action(description, *actions):
    # Display description & options; checks & returns input

    def get_checked_input(max_possible_input):
        flag = True
        while flag:
            input_to_check = input('\t')
            try:
                input_to_check = int(input_to_check)
                if (input_to_check > 0) & (input_to_check <= max_possible_input):
                    flag = False
                else:
                    error_message()
            except:
                error_message()
        return input_to_check

    print('\033[33m\n\t', description, '\033[0m')
    i = 0
    actions_list = []
    for action in actions:
        actions_list.append(action)
        i+=1
        print('\t\033[0m', i, ')\033[33m ', action, '\033[0m', sep='')
    choice_number = get_checked_input(i) - 1
    return actions_list[choice_number]

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

def error_message(*code_index):
    match code_index:
        case '0':
            print('VLAN does not exist in .xslx file')
        case _:
            print('\n\t\033[33m...you are doing something wrong, try again:\033[0m\n\t')

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
    def display_files():
        print('\t\033[33mHere are the files from my directory.\n\tSelect the .xlsx file in which you want to write the result:\n')
        files = os.listdir()
        file_index = 1
        actual_files = []
        ignored_files_list = ['backups', '.git', 'data', 'Maradeur-s-Map.py']
        for file_name in files:
            if (file_name in ignored_files_list):
                continue
            else:
                actual_files.append(file_name)
                current_string = '\033[0m \t' + str(file_index) + '\033[33m  ' + file_name 
                print(current_string)
                file_index += 1
        print('\t' + str(file_index) + '\033[33m  Refresh\033[33m')
        return actual_files
    all_files = display_files()
    file_digit = user_choosing(len(all_files))
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
    print('\n\t\033[33mChoose the VLAN you want to check:\n\033[0m')
    vlans_dict = VLAN_Administration.get_dict_of_vlans()
    arr_VLAN = []
    for vlan_name in vlans_dict:
        arr_VLAN.append(vlan_name)
    vlan_id = 1
    for vlan in arr_VLAN:
        current_string = '\033[0m\t' + str(vlan_id) + '\033[33m ' + arr_VLAN[vlan_id-1] 
        print(current_string)
        vlan_id+=1
    current_string = '\033[33m\t' + str(vlan_id) + ' Refresh\033[0m'
    print(current_string)
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
    new_xlsx_filename = input('\033[33mEnter a .xlsx filename: \033[0m')
    if not '.xslx' in new_xlsx_filename: new_xlsx_filename += '.xlsx'
    #new_xlsx_filename = "MTUCI Switches " + current_time + ".xlsx"
    change_log_filename = "changelog" + current_time + ".txt"
    return new_xlsx_filename, change_log_filename

def macAddCheckUtility(nmapDictionary, xlsx_file_path, sheetname_vlanname, new_changelogfile_filename): 
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
            except BaseException as e: print(e)
        index_of_mac_in_list+=1
    _read_xlsx_file.save(filename=new_xlsx_filename)

def ending(missingIp):
    print('\n\t\033[33mSuccess\033[0m')
    match len(missingIp):
        case 0:
            press_any()
        case _:
            print('\n\t\033[33m The following list contains information about those IP-addresses that are not in Google Sheet,\n\tbut the scanner detected devices behind IP. See "Missing Mac" in changelog\n')
            print(missingIp)

def scan_network_return_dict(network):
    nm = nmap.PortScanner()
    nmapDict = {}
    ip_address_list = []
    nets_dict = VLAN_Administration.get_dict_of_vlans()
    print('\033[33mScanning...')
    nmap_scan_dict = nm.scan(hosts=nets_dict[network], arguments='-n -sP -PE -PA21,23,80,3389')
    print('Done.\033[0m')
    ip_without_mac = []
    for ip in nmap_scan_dict['scan'].keys():
        ip_address_list.append(ip)
        try:
            macAdd = nmap_scan_dict['scan'][ip]['addresses']['mac']
        except:
            print('KeyError, scanned ip has no mac from nmap. The problem IP:', ip)
            ip_without_mac.append(ip)
            continue
        hostname = nmap_scan_dict['scan'][ip]['hostnames'][0]['name']
        nmapDict[ip] = macAdd
    return nmapDict, ip_without_mac

class VLAN_Administration :
    def get_dict_of_vlans():
        with open('data', 'rb') as f:
            vlan_dictionary = pickle.load(f)
        return vlan_dictionary
    def vlan_data_write(vlan_dict_to_write):
        with open('data', 'wb') as f:
            pickle.dump(vlan_dict_to_write, f)
    def display_vlans():
        vlan_dict = VLAN_Administration.get_dict_of_vlans()
        print('\033[33m\t Current VLANS:\n\t VLAN Name \t Network\033[0m')
        for vlan in vlan_dict:
            print('\t\033[0m', vlan, '\t', vlan_dict[vlan], '\033[0m')
    def add_vlan():
        def check_net_format(network_input):
            def check_max_val(arr):
                check_flag = True
                i = 0
                if len(arr) != 5: check_flag = False
                for val in arr:
                    try:
                        val = int(val)
                        if val > 255 : check_flag = False
                        if i == 4:
                            if (val > 32): check_flag = False
                        i+=1
                    except: check_flag = False
                return check_flag    
            flag = True
            while flag:
                network_input = network_input.strip()
                net_and_mask = network_input.split('/')
                match len(net_and_mask):
                    case 2:
                        arr_to_check = net_and_mask[0].split('.')
                        arr_to_check.append(net_and_mask[1])
                        if check_max_val(arr_to_check):
                            flag = False
                        else:
                            error_message()
                            network_input = input('\t')
                    case _:
                        error_message()
                        network_input = input('\t')
            return network_input
        new_vlanname = input('\033[33m\n\tEnter the VLAN name: \033[0m')
        new_network = input('\033[33m\tEnter the network address: \033[0m')
        new_network = check_net_format(new_network)
        dict_of_vlans = VLAN_Administration.get_dict_of_vlans()
        dict_of_vlans[new_vlanname] = new_network
        VLAN_Administration.vlan_data_write(dict_of_vlans)
        print('\033[33m\n\tSuccess!\033[0m\n')
        administrate_vlans()
    def delete_vlan():
        vlans_dict = VLAN_Administration.get_dict_of_vlans()
        vlan_to_remove = input('\t\033[33m Name the vlan you want to remove: \033[0m')
        vlans_dict.pop(vlan_to_remove, default)
        VLAN_Administration.vlan_data_write(vlans_dict)
        print('\t\033[33mVLAN \033[0m', vlan_to_remove, ' \033[33m has been removed.\033[0m\n')
        administrate_vlans()

def administrate_vlans():
    VLAN_Administration.display_vlans()
    action = choose_action('What do you want to do?', 'Add', 'Delete', 'Quit')
    match action:
        case 'Add':
            VLAN_Administration.add_vlan()
        case 'Delete':
            VLAN_Administration.delete_vlan()
        case 'Quit':
            start_menu()
            return 0

def start_menu():
    action = choose_action('Hey there', 'Scan my network', 'Manage my networks')
    match action:
        case 'Scan my network':
            xlsx_path = choose_xlsx_file()
            vlanname = choose_vlan()
            nmapDict, ip_without_mac = scan_network_return_dict(vlanname)
            edited_xlsx_filename, changelog_filename = get_name_for_files()
            correct_mac_list, missingIpDictionary= macAddCheckUtility(nmapDict, xlsx_path, vlanname, changelog_filename)
            changes_writer(xlsx_path, vlanname, correct_mac_list, edited_xlsx_filename)
            ending(missingIpDictionary)
            start_menu()
        case 'Manage my networks':
            administrate_vlans()

print_hello_words()
start_menu()