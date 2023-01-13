import os
import win32com.client
from pywinauto import Application
import keyboard
import psutil

# get somethings setup
stop = False
cwd = os.getcwd()
shell_link = win32com.client.Dispatch("WScript.Shell")
possible_starters = ['ctrl alt shift', 'ctrl shift', 'alt shift', 'ctrl', 'alt']


def load():
    # attempt to locate hotkeys file
    try:
        with open('hotkeys.txt', 'r') as f:
            # check every line
            for line in f:
                hotkey = line.split(',')[0]
                hotkey = hotkey.replace(" ", "+")
                # register hotkey
                keyboard.add_hotkey(hotkey, globals()[line.split(',')[1]], args=(line.split(',')[2],))

        print('loaded hotkeys at ' + cwd)

    # if it can't find the save then make new save file
    except FileNotFoundError:
        with open('hotkeys.txt', 'x'):
            print('no hotkeys file (did you rename the hotkeys.txt or move it?) new file created')


# create new hotkeys
def create():
    # get the starter keys from the user
    print('please insert the hotkey starters you would like to use')
    print('they should be in formatted like and in the order as follows')
    print('ctrl alt shift')
    print('shift cannot be used alone but you may use any other combination')
    print('if you would like to cancel creating a new hotkey at anytime type "cancel"')
    starters = input().lower()

    if starters == "cancel":
        print('hotkey canceled')
    # check that they are valid starter keys
    elif ("ctrl" in starters or "alt" in starters) and starters in possible_starters:
        # get the final alphanumeric key
        print('please enter a single alphanumeric key you would like to be used in combination with the starters')
        print('number pad and none alphanumeric keys hotkeys are currently in development')
        key = input().lower()
        override = False

        # check the hotkey doesn't already exist
        with open('hotkeys.txt', 'r') as f:
            contents = f.read()
            if starters + ' ' + key in contents:
                print('that hotkey combo already exists')
                print('do you want to override it? y/n')
                override = input()
                if override == 'y':
                    print('overriding')
                    override = True
                else:
                    key = 'cancel'

        if key == "cancel":
            print('hotkey canceled')

        elif key.isalnum() and len(key) == 1:
            # this is here as a structure for implementing other features in the future DO NOT DELETE IT
            # print('what do you want this hotkey to do')
            # print('to have it open or, if it is already open, focus an app type "open"')
            # action = input().lower()

            action = 'open'

            if action == "cancel":
                print('hotkey canceled')
            elif action == "open":
                print('please enter the path to the .exe for the app you would like it to open or focus')
                print('if you do not know how to find the path to the file please go to (fill me in).com')
                path = input()
                path = r'{}'.format(path)
                os.path.isfile(path)
                if path[-3:] == 'exe':
                    extension = 'exe'
                else:
                    extension = None

                if path == "cancel":
                    print('hotkey canceled')
                elif extension == 'exe':

                    print('to confirm your hotkey will be "' + starters + ' + ' + key + '" and it will ' + action +
                          ' the app at ' + path)
                    print('is this correct y/n')
                    answer = input().lower()

                    if answer == "y":
                        path = path.replace('\\', '/')
                        if not override:
                            with open('hotkeys.txt', 'a') as f:
                                f.seek(0, os.SEEK_END)
                                hotkey = starters + ' ' + key + ',run,' + path + '\n'
                                f.write(hotkey)
                        else:

                            with open("hotkeys.txt", "r+") as file:
                                lines = file.readlines()
                                for line_number, line in enumerate(lines):
                                    if starters + ' ' + key in line:
                                        keyboard.remove_hotkey(starters + ' + ' + key)
                                        lines[line_number] = starters + ' ' + key + ',run,' + path + '\n'
                                file.seek(0)
                                file.writelines(lines)
                                file.truncate()

                        print('hotkey saved')
                        starters = starters.replace(' ', '+')
                        keyboard.add_hotkey(starters + ' + ' + key, run, args=path)
                    else:
                        print('hotkey canceled')

                else:
                    print('invalid file, file must be a .exe file and include the file extension in the path')
        else:
            print('invalid input')
    else:
        print('invalid input')


# users +instructions
def commands():
    print('to create a new hotkey type "create"')
    print('to delete a hotkey type "delete"')
    print('to delete all hotkeys type "reset" this will also make sure')
    print('if you need help type "help"')
    print('please note that this app does not stop hotkeys in other apps from functioning')
    print('both the your focused app\'s hotkeys and the hotkeys you assign here will execute')


def run(path):
    found = False
    for proc in psutil.process_iter():
        try:
            base = os.path.normcase(proc.exe())
            base = os.path.basename(base)

            path_base = os.path.normcase(path)
            path_base = os.path.basename(path_base)
            if base == path_base.strip():
                found = True
                pid = proc.pid
                app = Application().connect(process=pid)
                app.top_window().set_focus()
                break
        except psutil.AccessDenied:
            pass

    if not found:
        print('opening')
        path = path.strip()
        os.startfile(path)


def remove_hotkey():
    print('what are the keys to the hotkey you would like to remove')
    print('use the format below and replace "key" with the final key in the hotkey')
    print('ctrl alt shift "key"')
    keys = input()
    with open(r'hotkeys.txt', 'r') as f:
        for line in f:
            if keys in line:
                print('are you sure you want to delete the following hotkey? y/n')
                print(keys + ' that ' + line.split(',')[1] + 's the program at ' + line.split(',')[2])
                confirm = input()
                if confirm == 'y':
                    f.write(line)
                    keyboard.remove_hotkey(keys)
                else:
                    continue


load()
commands()

while True:

    command = input().lower()
    if command == 'create':
        create()
    elif command == 'remove':
        print('unfinished')
    elif command == 'help':
        commands()
    elif command == 'reload':
        load()
        commands()
    elif command == 'reset':
        print('are you sure you want to delete ALL of you hotkeys? y/n')
        if input() == "y":
            open('hotkeys.txt', 'w').close()
            load()
            print('hotkeys reset')
        else:
            print('reset canceled')
    else:
        print('invalid command')
