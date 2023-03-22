from win32gui import FindWindow, CloseWindow, IsWindowVisible
from win32api import keybd_event, MapVirtualKey
import tkinter as tk
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import *
from time import sleep
import os
import winshell
import signal
import subprocess, re, os
from notifypy import Notify
import json
import ctypes
from win32com.client import Dispatch

path = r"H:\happy\hi\new.lnk"  # Path to be saved (shortcut)
target = r"D:\New folder\new.exe"  # The shortcut target file or folder
work_dir = r"D:\New folder"  # The parent folder of your file
name=str(os.path.basename(__file__))
banana=r'assets\banana.ico'

myappid = 'mycompany.myproduct.subproduct.version'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)



class Actions():

    def hwcode(Media):
        hwcode = MapVirtualKey(Media, 0)
        return hwcode

    def kill(app):
        notified=False
        found=0
        killed=0
        cmd = 'WMIC PROCESS get Caption,Processid'
        proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
        for line in proc.stdout:
            procces_info=str(line)
            if app in procces_info.lower():
                found+=1
                if notified==False:
                    notified=True
                    f=open('settings.json')
                    js=json.load(f)
                    if js['settings']['found_notification']['enabled']==1:
                        notification = Notify()
                        notification.title = 'Advertisement'
                        notification.icon = banana
                        notification.application_name = name
                        notification.message = 'Add found in: '+app
                        notification.send()
                if re.findall(r'[0-9]{3,6}', procces_info):
                    pid=re.findall(r'[0-9]{3,6}', procces_info)[0]
                    print('attempting to kill: '+pid)
                    try:
                        os.kill(int(pid),signal.SIGTERM)
                    except:
                        print('failed to kill: '+pid)
                    else:
                        print('killed: '+pid)
                        killed+=1
        return found, killed

    def resume(reply, app):
        keybd_event(0xB3, Actions.hwcode(0xB3)) #play
        keybd_event(0xB0, Actions.hwcode(0xB0)) #next
        f=open('settings.json')
        js=json.load(f)
        if js['settings']['debug_notification']['enabled']==1:
            notification = Notify()
            notification.title = 'Add skipped'
            notification.icon = banana
            notification.application_name = name
            notification.message = 'Succesfully resumed '+app+', killed: '+str(reply[1])+' of '+str(reply[0])+' '+app+' processes.'
            notification.send()

    def escape(reason):
        notification = Notify()
        notification.title = 'Closing Banana.'
        notification.icon = banana
        notification.application_name = name
        notification.message = reason
        notification.send()
        quit('Exiting program')

    def apology(app):
        notification = Notify()
        notification.title = 'Application not found'
        notification.icon = banana
        notification.application_name = name
        notification.message = app+' could not be opened, closing banana.'
        notification.send()
        quit('Exiting program, apology.')

    def error():
        notification = Notify()
        notification.title = 'Whoops'
        notification.icon = banana
        notification.application_name = name
        notification.message = 'Something unknown went wrong, closing banana.'
        notification.send()
        quit('Exiting program, error.')

    def open_menu(self):
        window=GUI()
        window.CreateWindow()

class app(Actions):
    def main(self):
        print('running...')
        f=open('settings.json')
        js=json.load(f)
        if js['settings']['setting_startup']['enabled']==1 or not os.path.exists('assets\Spotify.lnk'):
            Actions.open_menu(self)
        if js['settings']['start_notification']['enabled']==1:
            notification = Notify()
            notification.title = 'Banana started'
            notification.icon = banana
            notification.application_name = name
            notification.message = 'Listening for window changes.'
            notification.send()

        while True:
            sleep(.5)
            for add in js['blacklist']:
                foobar1 = FindWindow(None, add)
                if foobar1:
                    reply=Actions.kill('spotify')
                    os.startfile('assets\Spotify.lnk')
                    for tries in range(0, 10):
                        spotify = FindWindow(None, 'Spotify Free')
                        if spotify:
                            CloseWindow(FindWindow(None, 'Spotify Free'))
                            break
                        if tries==10:
                            Actions.apology('spotify')
                        sleep(.5)
                    Actions.resume(reply, 'spotify')

            foobar2 = FindWindow(None, 'Spotify Free')
            if foobar2:
                if IsWindowVisible(foobar2) == False:
                    reply=Actions.kill('spotify')
                    os.startfile('assets\Spotify.lnk')
                    for tries in range(0, 10):
                        spotify = FindWindow(None, 'Spotify Free')
                        if spotify:
                            CloseWindow(FindWindow(None, 'Spotify Free'))
                            break
                        if tries==10:
                            Actions.apology('spotify')
                        sleep(.5)
                    Actions.resume(reply, 'spotify')
class events():
    def start(self, master):
        master.destroy()

    def refresh(self, master):
        master.destroy()
        GUI.CreateWindow(self)

    def getPath(self):
        with open("settings.json", "r") as jsonFile:
            data = json.load(jsonFile)
            data= data["settings"]['path']
        return data
    
    def getKey(self):
        with open("settings.json", "r") as jsonFile:
            data = json.load(jsonFile)
            data= data["settings"]['force_skip_key']
        return data

    def setNewPath(self, master):
        master.filename = filedialog.askopenfilename(initialdir = "/",title = "Select Spotify.exe",filetypes = (("exe files","*.exe"),("all files","*.*")))
        with open("settings.json", "r") as jsonFile:
            data = json.load(jsonFile)
            data["settings"]['path'] = master.filename
        with open("settings.json", "w") as jsonFile:
            json.dump(data, jsonFile,indent = 4)
        if not os.path.exists('assets\Spotify.lnk'):
            open('assets\Spotify.lnk', 'x')
        path = "assets\Spotify.lnk"  # Path to be saved (shortcut)
        s=str(master.filename).split('/')[-1]
        target = str(master.filename) # The shortcut target file or folder
        work_dir = str(master.filename).replace(s,'')  # The parent folder of your file
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(path)
        shortcut.Targetpath = target
        shortcut.WorkingDirectory = work_dir
        shortcut.save()
        self.refresh(master)
        return master.filename

    def changeNotifications(self, js, type):
        print(js["settings"][type],js["settings"][type]['enabled'])
        with open("settings.json", "r") as jsonFile:
            data = json.load(jsonFile)
        if data["settings"][type]['enabled']==1:
            data["settings"][type]['enabled']= 0
        else:
            data["settings"][type]['enabled']= 1
        with open("settings.json", "w") as jsonFile:
            json.dump(data, jsonFile,indent = 4)
        print(js["settings"][type],js["settings"][type]['enabled'])

# Create Frame
class GUI():
    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            quit()

    def CreateWindow(self):

        f=open('settings.json')
        js=json.load(f)
        remote=events()
        master = Tk()

        master.geometry("600x300")
        master.title('Spotify with Banana.py')
        master.iconphoto(False, PhotoImage(file='assets/banana.ico'))
        master.resizable(width=False, height=False)

        bg = PhotoImage(file = "assets/banana.png")
        label1 = Label( master, image = bg)
        label1.place(x = -10, y = -80)

        s=Style()
        s.configure('My.TFrame', background='black', highlightbackground="yellow",highlightthickness=5)
        s.configure('My2.TFrame', background='yellow', highlightbackground="black",highlightthickness=5)

        main1 = Frame(master, style='My2.TFrame')
        main1.pack(padx=50, pady=50)
        main = Frame(main1, style='My.TFrame')
        main.pack(padx=1, pady=1)
        textvariable=remote.getPath()
        keyvariable=remote.getKey()
        if textvariable=="":
            textvariable="No path selected"
        if keyvariable=="":
            keyvariable="No key bound"
        cframe = Frame(main, style='My.TFrame')
        cframe.pack(pady=10, padx=10, side=TOP)

        a=tk.Checkbutton(cframe,
            text = js['settings']['start_notification']['title'],
            bg='black', fg='grey',
            variable=IntVar(value=js['settings']['start_notification']['enabled']),
            command=lambda:remote.changeNotifications(js, 'start_notification'),
            highlightbackground="yellow",
        )
        b=tk.Checkbutton(cframe,
            text = js['settings']['found_notification']['title'],
            bg='black', fg='grey',
            variable=IntVar(value=js['settings']['found_notification']['enabled']),
            command=lambda:remote.changeNotifications(js, 'found_notification'),
            highlightbackground="yellow",
        )
        c=tk.Checkbutton(cframe,
            text = js['settings']['debug_notification']['title'],
            bg='black', fg='grey',
            variable=IntVar(value=js['settings']['debug_notification']['enabled']),
            command=lambda:remote.changeNotifications(js, 'debug_notification'),
            highlightbackground="yellow",
        )
        d=tk.Checkbutton(cframe,
            text = js['settings']['setting_startup']['title'],
            bg='black', fg='grey',
            variable=IntVar(value=js['settings']['setting_startup']['enabled']),
            command=lambda:remote.changeNotifications(js, 'setting_startup'),
            highlightbackground="yellow",
        )
        e=tk.Checkbutton(cframe,
            text = js['settings']['start_on_launch']['title'],
            bg='black', fg='grey',
            variable=IntVar(value=js['settings']['start_on_launch']['enabled']),
            command=lambda:remote.changeNotifications(js, 'start_on_launch'),
            highlightbackground="yellow",
        )
        f2=tk.Checkbutton(cframe,
            text = js['settings']['blacklist_notifications']['title'],
            bg='black', fg='grey',
            variable=IntVar(value=js['settings']['blacklist_notifications']['enabled']),
            command=lambda:remote.changeNotifications(js, 'blacklist_notifications'),
            highlightbackground="yellow",
        )
        if js['settings']['start_notification']['enabled']==1:
            a.select()
        if js['settings']['found_notification']['enabled']==1:
            b.select()
        if js['settings']['debug_notification']['enabled']==1:
            c.select()
        if js['settings']['setting_startup']['enabled']==1:
            d.select()
        if js['settings']['start_on_launch']['enabled']==1:
            e.select()
        if js['settings']['blacklist_notifications']['enabled']==1:
            f2.select()
        a.pack()
        b.pack()
        c.pack()
        d.pack()
        # e.pack()
        # f2.pack()

        setframe = Frame(main, style='My.TFrame')
        setframe.pack(pady=1, padx=1,side=TOP)
        if keyvariable=="No key bound":
            label = tk.Label(setframe, text = 'Force skip key: [NONE]', fg='yellow', background='black')
            label.pack(padx=5,pady=5, side=RIGHT)
            submit2 = tk.Button(setframe,text ='bind key', bg='black', fg='orange')
            submit2.pack(pady=5,side=LEFT)
        else:
            label = tk.Label(setframe, text = 'Force skip key: ['+keyvariable+']', fg='yellow', background='black')
            label.pack(padx=5,pady=5, side=RIGHT)
            submit2 = tk.Button(setframe,text ='change key', bg='black', fg='lime')
            submit2.pack(pady=5,side=LEFT)

        btn = tk.Button(main,text ='set path', bg='black', fg='yellow',command =lambda:remote.setNewPath(master))
        btn.pack(padx=5,pady=5,side=LEFT)
        if textvariable=="No path selected":
            label = tk.Label(main, text = textvariable, foreground='red', background='black')
            label.pack(padx=5,pady=5, side=RIGHT)
        else:
            label = tk.Label(main, text = textvariable, foreground='yellow', background='black')
            label.pack(padx=5,pady=5, side=RIGHT)
        if textvariable!="No path selected":
            if 'spotify' in str(textvariable).lower() and '.exe' in str(textvariable).lower():
                submit = tk.Button(main,text ='start banana', bg='black', fg='lime',command =lambda: remote.start(master))
                submit.pack(pady=5,side=RIGHT)
            elif not '.exe' in str(textvariable).lower():
                submit = tk.Button(main,text ='path to file needs to be a .exe', bg='black', fg='grey')
                submit.pack(pady=5,side=RIGHT)
            else:
                submit = tk.Button(main,text ='Invalid path, start anyway?', bg='black', fg='orange',command =lambda: remote.start(master))
                submit.pack(pady=5,side=RIGHT)
        elif textvariable=="No path selected":
            submit = tk.Button(main,text ='start banana', bg='black', fg='grey')
            submit.pack(pady=5,side=RIGHT)

        master.protocol("WM_DELETE_WINDOW", lambda:GUI.on_closing(self))
        mainloop()

try:
    main = app()
    main.main()
except KeyboardInterrupt:
    Actions.escape('KeyboardInterrupt. Closing.')
except Exception as e:
    Actions.escape('Banana Interrupted. Exception E.')
except:
    Actions.escape('Banana Closed.')

