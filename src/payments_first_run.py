# -*- coding: utf-8 -*-
"""
Created on Tue Aug 20 14:54:51 2019

@author: v.shkaberda
"""
from _version import upd_path
from contextlib import suppress
from tkinter import Tk, messagebox
from win32com.client import Dispatch
import getpass, os, zipfile, zlib

SOURCE = zlib.decompress(upd_path).decode()
TARGET = os.path.join('C:\\Users', getpass.getuser(), 'AppData\\Local')
REFERENCES = os.path.join('C:\\Users',
                          getpass.getuser(),
                          'AppData\\Roaming\\Microsoft\\Windows\\Payments')
DESKTOP = os.path.join('C:\\Users', getpass.getuser(), 'Desktop')
WDIR = os.path.join(TARGET, 'Payments')
TARGETFILE = os.path.join(WDIR, 'payments_checker.exe')
ICONFILE = os.path.join(WDIR, 'resources\\payment.ico')

class SuccessMsg(Tk):
    """ Raise an error when user doesn't have permission to work with app.
    """
    def __init__(self):
        super().__init__()
        self.withdraw()  # Do not show main window
        messagebox.showinfo(
            'Заявки на платежи',
            'Установка завершена.\n'
            'На рабочем столе создан ярлык для запуска.'
        )
        self.destroy()

def create_shortcut(path, target='', wDir='', icon=''):
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = wDir
    shortcut.Description = 'Заявки на оплату'
    if icon:
        shortcut.IconLocation = icon
    shortcut.save()

def main():
    print('Выполняется начальная установка и создание ярлыков...')
    # extract actual version of app
    with zipfile.ZipFile(os.path.join(SOURCE, 'Payments.zip'), 'r') as zip_ref:
        zip_ref.extractall(TARGET)
    # create shortcuts in windows start menu and on desktop
    with suppress(FileExistsError):
        os.mkdir(REFERENCES)
    for path in (REFERENCES, DESKTOP):
        create_shortcut(os.path.join(path, 'Заявки на оплату.lnk'),
                        TARGETFILE,
                        WDIR,
                        ICONFILE)
    SuccessMsg()

if __name__ == '__main__':
    main()