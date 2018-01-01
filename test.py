#!/usr/bin/python

import sys
from appJar import gui

def main():
    # _winreg is renamed to winreg in Python 3.x
    if sys.version_info >= (3, 0):
        from winreg import CreateKey, SetValueEx, CloseKey, HKEY_CURRENT_USER, KEY_WRITE, KEY_READ, REG_SZ
    else:
        from _winreg import CreateKey, SetValueEx, CloseKey, HKEY_CURRENT_USER, KEY_WRITE, KEY_READ, REG_SZ
    
    # NOTE: taken from https://stackoverflow.com/questions/15128225/python-script-to-read-and-write-a-path-to-registry
    REG_PATH = r"SOFTWARE\fmpof\Settings"
    def set_reg(name, value):
        try:
            CreateKey(HKEY_CURRENT_USER, REG_PATH)
            registry_key = OpenKey(HKEY_CURRENT_USER, REG_PATH, 0, KEY_WRITE)
            SetValueEx(registry_key, name, 0, REG_SZ, value)
            CloseKey(registry_key)
            return True
        except WindowsError:
            return False
    def get_reg(name):
        try:
            registry_key = OpenKey(HKEY_CURRENT_USER, REG_PATH, 0, KEY_READ)
            value, regtype = QueryValueEx(registry_key, name)
            CloseKey(registry_key)
            return value
        except WindowsError:
            return None

    app = gui()
    app.addLabel("Title", "Forgot My Passwords of Office Files")
    app.setLabelBg("Title", "gray")
    app.addFileEntry("dbfile")
    app.addLabelSecretEntry("pass")
    app.go()

if __name__ == "__main__":
    main()
