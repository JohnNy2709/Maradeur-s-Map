from email.policy import default
import pickle
import os
import ctypes
kernel32 = ctypes.windll.kernel32
kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)
def clear_screen():
    os.system('cls')

def error_message():
    print('\n\t\033[33m...you are doing something wrong, try again:\033[0m\n\t')

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
        new_vlanname = input('\033[33m\n\tEnter the VLAN name: \033[0m')
        new_network = input('\033[33m\tEnter the network address: \033[0m')
        dict_of_vlans = VLAN_Administration.get_dict_of_vlans()
        dict_of_vlans[new_vlanname] = new_network
        VLAN_Administration.vlan_data_write(dict_of_vlans)
        print('\033[33muccess!\033[0m')
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
        case 'Back':
            return 'back'

administrate_vlans()