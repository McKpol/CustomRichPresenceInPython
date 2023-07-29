import customtkinter
import threading
import pystray
import PIL.Image
from pypresence import Presence
import os
import time
import getpass
from win32com.client import Dispatch

USER_NAME = getpass.getuser()

if not os.path.exists("settings.txt"):
    with open("settings.txt", 'w') as plik:
        plik.write("False\nFalse\ndark")

if not os.path.exists("save.txt"):
    with open("save.txt", 'w') as plik:
        plik.write("1122815026697355385\nHey, i am very cool guy!\nI like typing random thing on chat discord!\nhttps://assets.stickpng.com/images/62b2265e038aad4d3ed7ca4b.png\nDiscord Logo\nhttps://c.tenor.com/TgKK6YKNkm0AAAAi/verified-verificado.gif\nVerified\nTrue\n1\n10\nThis Custom Rich Presence\nhttps://github.com/McKpol/CustomRichPresenceInPython\nHe's YouTube\nhttps://www.youtube.com/@McKLeniwiec")

def add_to_startup(file_path=""):
    print("Utworzono")
    if file_path == "":
        file_path = os.path.dirname(os.path.realpath(__file__))
    bat_path = r'C:\Users\%s\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup' % USER_NAME
    path = f"{bat_path}\CustomRichPresenceDiscord.lnk" # Path to be saved (shortcut)
    target = f"{file_path}\main.exe"  # The shortcut target file or folder
    work_dir = f"{file_path}"  # The parent folder of your file

    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = work_dir
    shortcut.save()

def delete_to_startup():
    print("Czekanie na usunięcie")
    bat_path = r'C:\Users\%s\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup' % USER_NAME
    path = (bat_path + r'\\' + r"CustomRichPresenceDiscord.lnk")
    if os.path.exists(path):
        os.remove(path)
        print("Plik został usunięty.")
    else:
        print("Plik nie istnieje.")
    print(f"{bat_path}\\CRPD.ink")

def getsaveline(line):
    try:
        with open('save.txt', 'r') as file:
            lines = file.readlines()

        line_number = line
        if line_number <= len(lines):
            selected_line = lines[line_number - 1]
            return(selected_line.strip())
        file.close()
    except FileNotFoundError: print("Plik nie znaleziony")

AppID = getsaveline(1)
Details = getsaveline(2)
State = getsaveline(3)
LargeImageLink = getsaveline(4)
LargeImageDescription = getsaveline(5)
SmallImageLink = getsaveline(6)
SmallImageDescription = getsaveline(7)
PartyShow = getsaveline(8)
PartyNumber = getsaveline(9)
PartySlots = getsaveline(10)
Button1Name = getsaveline(11)
Button1Link = getsaveline(12)
Button2Name = getsaveline(13)
Button2Link = getsaveline(14)

def getsettingsline(line):
    try:
        with open('settings.txt', 'r') as file:
            lines = file.readlines()

        line_number = line
        if line_number <= len(lines):
           selected_line = lines[line_number - 1]
           return(selected_line.strip())
        file.close()
    except FileNotFoundError: print("Plik nie znaleziony")

startup = getsettingsline(1)
minimalize = getsettingsline(2)
apperancemode = getsettingsline(3)

if getsettingsline(1) == None:
    startup = "False"
if getsettingsline(2) == None:
    minimalize = "False"
if getsettingsline(3) == None:
    apperancemode = "dark"

class Settings(customtkinter.CTkToplevel):
    def __init__(self):
        super().__init__()
        self.geometry("400x300")
        self.grid_columnconfigure(1, weight=1)
        self.title("Settings")

        def savesettings():
            print("Zapisano ustawienia")
            if self.checkboxstartup.get() == "True":
                add_to_startup()
            else: 
                delete_to_startup()
            startup = self.checkboxstartup.get()
            minimalize = self.checkboxminimalize.get()
            appearancemode = self.appearance_mode_optionemenu.get()
            with open('settings.txt', 'w')  as file:
                file.write(f"{startup}\n{minimalize}\n{appearancemode}")
                file.close()
        self.label = customtkinter.CTkLabel(self, text="Setting", font=customtkinter.CTkFont(family="Arial", size=20, weight="bold"))
        self.label.grid(row=0, column=1, padx=0, pady=10)
        self.checkboxstartup = customtkinter.CTkCheckBox(self, text="Run on startup Windows", variable=customtkinter.StringVar(value=startup), onvalue="True", offvalue="False")
        self.checkboxstartup.grid(row=1, column=1, padx=0, pady=5)
        self.checkboxminimalize = customtkinter.CTkCheckBox(self, text="Start program minimalize", variable=customtkinter.StringVar(value=minimalize), onvalue="True", offvalue="False")
        self.checkboxminimalize.grid(row=2, column=1, padx=0, pady=5)
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self, values=["light", "dark", "system"],command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.set(apperancemode)
        self.appearance_mode_optionemenu.grid(row=3, column=1, padx=0, pady=5)
        self.buttonsave = customtkinter.CTkButton(self, text="Save", command=savesettings)
        self.buttonsave.grid(row=4, column=1, padx=0, pady=5)
    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)



class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        def getsettingsline(line):
            try:
                with open('settings', 'r') as file:
                    lines = file.readlines()

                line_number = line
                if line_number <= len(lines):
                   selected_line = lines[line_number - 1]
                   return(selected_line.strip())
                file.close()
            except FileNotFoundError: print("Plik nie znaleziony")
        self.iconbitmap("icon.ico")
        self.title("Custom Rich Presence by McKpl")
        self.geometry(f"{1020}x{310}")
        self.minsize(1020, 310)
        self.maxsize(1020, 310)
        self.grid_columnconfigure((0, 6), weight=1)
        self._set_appearance_mode(apperancemode)

        self.labelMenuText = customtkinter.CTkLabel(self, text="Custom Rich Presence Discord", font=customtkinter.CTkFont(family="Arial", size=35, weight="bold"))
        self.labelMenuText.grid(row=0, column=1, padx=0, pady=0, sticky="ew", columnspan=6)

        self.botFrame = customtkinter.CTkFrame(self)
        self.botFrame.grid(row=1, column=1, padx=10, pady=0, sticky="nsw", columnspan=6)

        self.textapp = customtkinter.CTkLabel(self.botFrame, text="APP ID", font=customtkinter.CTkFont(family="Arial", size=15, weight="bold"))
        self.textapp.grid(row=0, column=0, padx=10, pady=2)

        self.entryapp = customtkinter.CTkEntry(self.botFrame, placeholder_text="APP_ID", width=200)
        self.entryapp.insert(0, AppID)
        self.entryapp.grid(row=0, column=1, padx=2, pady=2)

        self.info_frame = customtkinter.CTkFrame(self)
        self.info_frame.grid(row=2, column=1, padx=10, pady=10, sticky="nsw")

        self.labeltext = customtkinter.CTkLabel(self.info_frame, text="Text", font=customtkinter.CTkFont(family="Arial", size=20, weight="bold"))
        self.labeltext.grid(row=0, column=0, padx=5, pady=5, sticky="we")

        self.entrydetails = customtkinter.CTkEntry(self.info_frame, placeholder_text="Details")
        self.entrydetails.insert(0, Details)
        self.entrydetails.grid(row=1, column=0, padx=5, pady=5)
        
        self.entrystate = customtkinter.CTkEntry(self.info_frame, placeholder_text="State")
        self.entrystate.insert(0, State)
        self.entrystate.grid(row=2, column=0, padx=5, pady=5)

        self.large_image = customtkinter.CTkFrame(self)
        self.large_image.grid(row=2, column=2 , padx=10, pady=10, sticky="nsw")

        self.labelLargeImage = customtkinter.CTkLabel(self.large_image, text="Large Image", font=customtkinter.CTkFont(family="Arial", size=20, weight="bold"))
        self.labelLargeImage.grid(row=0, column=0, padx=5, pady=5, sticky="we")

        self.entryLinkLargeImage = customtkinter.CTkEntry(self.large_image, placeholder_text="Link")
        self.entryLinkLargeImage.insert(0, LargeImageLink)
        self.entryLinkLargeImage.grid(row=1, column=0, padx=5, pady=5)
        
        self.entryDescriptionLargeImage = customtkinter.CTkEntry(self.large_image, placeholder_text="Description")
        self.entryDescriptionLargeImage.insert(0, LargeImageDescription)
        self.entryDescriptionLargeImage.grid(row=2, column=0, padx=5, pady=5)

        self.small_image = customtkinter.CTkFrame(self)
        self.small_image.grid(row=2, column=3 , padx=10, pady=10, sticky="nsw")

        self.labelSmallIlamge = customtkinter.CTkLabel(self.small_image, text="Small Image", font=customtkinter.CTkFont(family="Arial", size=20, weight="bold"))
        self.labelSmallIlamge.grid(row=0, column=0, padx=5, pady=5, sticky="we")

        self.entryLinkSmallImage = customtkinter.CTkEntry(self.small_image, placeholder_text="Link")
        self.entryLinkSmallImage.insert(0, SmallImageLink)
        self.entryLinkSmallImage.grid(row=1, column=0, padx=5, pady=5)
        
        self.entryDescriptionSmallImage = customtkinter.CTkEntry(self.small_image, placeholder_text="Description")
        self.entryDescriptionSmallImage.insert(0, SmallImageDescription)
        self.entryDescriptionSmallImage.grid(row=2, column=0, padx=5, pady=5)

        self.party = customtkinter.CTkFrame(self)
        self.party.grid(row=2, column=4, padx=10, pady=10, sticky="nsw")

        self.labelParty = customtkinter.CTkLabel(self.party, text="Party", font=customtkinter.CTkFont(family="Arial", size=20, weight="bold"))
        self.labelParty.grid(row=0, column=0, padx=5, pady=5, sticky="we")

        self.switch_var = customtkinter.StringVar(value=PartyShow)
        self.switchParty = customtkinter.CTkSwitch(self.party, text="Show", variable=self.switch_var, onvalue="True", offvalue="False")
        self.switchParty.grid(row=1, column=0, padx=5, pady=5)
        
        self.entryPeople = customtkinter.CTkEntry(self.party, placeholder_text="Number of people")
        self.entryPeople.insert(0, PartyNumber)
        self.entryPeople.grid(row=2, column=0, padx=5, pady=5)
        
        self.entrySlots = customtkinter.CTkEntry(self.party, placeholder_text="Number of slots")
        self.entrySlots.insert(0, PartySlots)
        self.entrySlots.grid(row=3, column=0, padx=5, pady=5)

        self.Button1 = customtkinter.CTkFrame(self)
        self.Button1.grid(row=2, column=5, padx=10, pady=10, sticky="nsw")

        self.labelButton1 = customtkinter.CTkLabel(self.Button1, text="Button 1", font=customtkinter.CTkFont(family="Arial", size=20, weight="bold"))
        self.labelButton1.grid(row=1, column=0, sticky="we")

        self.entryname1 = customtkinter.CTkEntry(self.Button1, placeholder_text="Entry Name")
        self.entryname1.insert(0, Button1Name)
        self.entryname1.grid(row=2, column=0, padx=5, pady=5)

        self.entrylink1 = customtkinter.CTkEntry(self.Button1, placeholder_text="Entry Link")
        self.entrylink1.insert(0, Button1Link)
        self.entrylink1.grid(row=3, column=0, padx=5, pady=5)

        self.Button2 = customtkinter.CTkFrame(self)
        self.Button2.grid(row=2, column=6, padx=10, pady=10, sticky="nsw")

        self.labelButton2 = customtkinter.CTkLabel(self.Button2, text="Button 2", font=customtkinter.CTkFont(family="Arial", size=20, weight="bold"))
        self.labelButton2.grid(row=1, column=0, sticky="we")

        self.entryname2 = customtkinter.CTkEntry(self.Button2, placeholder_text="Entry Name")
        self.entryname2.insert(0, Button2Name)
        self.entryname2.grid(row=2, column=0, padx=5, pady=5)

        self.entrylink2 = customtkinter.CTkEntry(self.Button2, placeholder_text="Entry Link")
        self.entrylink2.insert(0, Button2Link)
        self.entrylink2.grid(row=3, column=0, padx=5, pady=5)
        
        self.buttonstart = customtkinter.CTkButton(self, text="Start", command=self.openRPC)
        self.buttonstart.grid(row=3, column=1, pady=2, padx=(10, 1), sticky="ew", columnspan=3)

        self.buttonsave = customtkinter.CTkButton(self, text="Save", command=self.saveRPC)
        self.buttonsave.grid(row=3, column=4, pady=2, padx=(1, 10), sticky="ew", columnspan=3)

        self.buttonhide = customtkinter.CTkButton(self, text="Hide", command=self.hide)
        self.buttonhide.grid(row=4, column=1, pady=2, padx=(10, 1), sticky="ew", columnspan=3)

        self.buttonsettings = customtkinter.CTkButton(self, text="Settings", command=self.open_settings)
        self.buttonsettings.grid(row=4, column=4, pady=2, padx=(1, 10), sticky="ew", columnspan=3)

        if minimalize == "True":
            self.hide()

        self.toplevel_window = None

    def open_settings(self):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = Settings()  # create window if its None or destroyed
        else:
            self.toplevel_window.focus()  # if window exists focus it
        self.toplevel_window.focus()  # if window exists focus it

    def saveRPC(self):

        print("Zapiswyanie Pliku")
        AppID = self.entryapp.get()
        Details = self.entrydetails.get()
        State = self.entrystate.get()
        LargeImageLink = self.entryLinkLargeImage.get()
        LargeImageDescription = self.entryDescriptionLargeImage.get()
        SmallImageLink = self.entryLinkSmallImage.get()
        SmallImageDescription = self.entryDescriptionSmallImage.get()
        PartyShow = self.switchParty.get()
        PartyNumber = self.entryPeople.get()
        PartySlots = self.entrySlots.get()
        Button1Name = self.entryname1.get()
        Button1Link = self.entrylink1.get()
        Button2Name = self.entryname2.get()
        Button2Link = self.entrylink2.get()

        with open('save.txt', 'w')  as file:
            file.write(f"{AppID}\n{Details}\n{State}\n{LargeImageLink}\n{LargeImageDescription}\n{SmallImageLink}\n{SmallImageDescription}\n{PartyShow}\n{PartyNumber}\n{PartySlots}\n{Button1Name}\n{Button1Link}\n{Button2Name}\n{Button2Link}")
            file.close()
        print("Zapisano!")
    
    def updateRPC(self):
        print("Uruchamianie....")
        timestamp = time.time()
        self.RPCid = Presence(int(AppID))
        try: self.RPCid.connect()
        except Exception as e: 
            print(f"Złe AppID: {e}")
            return 
        try:
            self.RPCid.update(
            state=State,
            details=Details,
            large_image=LargeImageLink,
            large_text=LargeImageDescription,
            small_image=SmallImageLink,
            small_text=SmallImageDescription,
            start=timestamp,
            party_size=[int(PartyNumber),int(PartySlots)],
            buttons=[{"label": Button1Name, "url": Button1Link},{"label": Button2Name, "url": Button2Link}]
            )
            print("Refresh")
        except Exception as e:
            print(f"Złe ustawienie: {e}")
        return

    def closeRPC(self):
        print("Zamykanie RPC")
        self.RPCid.close()


    def selfcloseRPC(self):
        self.closeRPC()
        self.buttonstart.configure(text="Start", command=self.openRPC)

    
    def openRPC(self):
        print ("Zatrzymywanie RPC")
        try: 
            thread.join()
        except Exception: print("Nie udało się")
        print("Zapiswyanie Pliku")
        AppID = self.entryapp.get()
        Details = self.entrydetails.get()
        State = self.entrystate.get()
        LargeImageLink = self.entryLinkLargeImage.get()
        LargeImageDescription = self.entryDescriptionLargeImage.get()
        SmallImageLink = self.entryLinkSmallImage.get()
        SmallImageDescription = self.entryDescriptionSmallImage.get()
        PartyShow = self.switchParty.get()
        PartyNumber = self.entryPeople.get()
        PartySlots = self.entrySlots.get()
        Button1Name = self.entryname1.get()
        Button1Link = self.entrylink1.get()
        Button2Name = self.entryname2.get()
        Button2Link = self.entrylink2.get()

        with open('save.txt', 'w')  as file:
            file.write(f"{AppID}\n{Details}\n{State}\n{LargeImageLink}\n{LargeImageDescription}\n{SmallImageLink}\n{SmallImageDescription}\n{PartyShow}\n{PartyNumber}\n{PartySlots}\n{Button1Name}\n{Button1Link}\n{Button2Name}\n{Button2Link}")
            file.close()
        print("Zapisano!")
        print("Uruchamianie...")
        thread = threading.Thread(target=self.updateRPC())
        thread.start
        self.buttonstart.configure(text="Stop", command=self.selfcloseRPC)

    def hide(self):
        self.withdraw()
        image = PIL.Image.open("icon.png")

        def on_clicked():
            self.deiconify()
            icon.stop()
        def exit():
            icon.stop()
            os._exit(0)

        icon = pystray.Icon("Icon", image, menu=pystray.Menu(
        pystray.MenuItem("Show", on_clicked),
        pystray.MenuItem("Exit", exit),
        ))

        icon.run_detached()

app = App()
app.mainloop()