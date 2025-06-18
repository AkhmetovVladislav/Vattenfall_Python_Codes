import customtkinter
import tkinter
import os
from CTkMessagebox import CTkMessagebox
import Allokation_ANB_Tag_VES_VREES
import VES_APM_RLM_SLP
import VES_SLP_Strom
import VREES_Fahrplandaten_Strom
import VREES_Gasanlagen

# Bestimme den Benutzernamen
benutzername = os.getlogin()

customtkinter.set_appearance_mode("light") # die Farbe des Fensters
customtkinter.set_default_color_theme("dark-blue") # die Farbe der Buttons

app = customtkinter.CTk()
app.resizable(width=False, height=False)

app.title("Excel-Ausfüller") # Titel des Fensters
app.geometry("400x300") # Größe des Fensters in pixels

def show_info():
    # Default messagebox for showing some information
    CTkMessagebox(title="Info", message="Hol Dir mal einen Kaffee. Die Aktualisierung wird etwas dauern. Viel Spaß!")

def show_checkmark():
    # Show some positive message with the checkmark icon
    CTkMessagebox(title="Info", message="Alle ausgewählten Dateien wurden erfolgreich aktualisiert",
                  icon="check", option_1="Perfekt")

def show_error():
    # Show some error message
    CTkMessagebox(title="Fehler", message="Keine Datei wurde ausgewählt", icon="cancel")

def ausführen():
    # Liste der Funktionen und zugehörigen Variablen
    funktionen = [
        (VREES_Gasanlagen.ausführen, check_button_1_var, "VREES_Gasanlagen.xlsx"),
        (VREES_Fahrplandaten_Strom.ausführen, check_button_2_var, "VREES_Fahrplandaten_Strom.xlsx"),
        (VES_SLP_Strom.ausführen, check_button_3_var, "VES_SLP_Strom.xlsx"),
        (VES_APM_RLM_SLP.ausführen, check_button_4_var, "VES_APM_RLM_SLP.xlsx"),
        (Allokation_ANB_Tag_VES_VREES.ausführen, check_button_6_var, "Allokation_ANB_Tag_VES_VREES.xlsx")
    ]
    
    # Überprüfe, ob mindestens ein Checkbutton ausgewählt ist
    mindestens_ein_checkbutton = any(var.get() == "on" for _, var, _ in funktionen)
    
    if mindestens_ein_checkbutton == False:
        show_error()  # Ansonsten, falls kein Checkbutton ausgewählt wurde, rufe die Funktion show_error() auf

    for funktion, var, datei in funktionen:
        if var.get() == "on":
            try:
                funktion()
            except Exception as e:
                CTkMessagebox(title="Fehler", message=f"Bei der Datei '{datei}' ist ein Fehler aufgetreten", icon="cancel")

    if mindestens_ein_checkbutton:           
        show_checkmark()

label = customtkinter.CTkLabel(master=app,
                               text="Wähle mal die Excel-Dateien aus, die aktualisiert werden sollen",
                               width=120,
                               height=25,
                               fg_color=("white"),
                               corner_radius=8)
label.place(relx=0.05, rely=0.05)

frame_height = 0.175

check_button_1_var = customtkinter.StringVar(value="off")
check_button_1 = customtkinter.CTkCheckBox(master=app, text="VREES_Gasanlagen.xlsx", variable=check_button_1_var, onvalue="on", offvalue="off")
check_button_1.place(relx=0.05, rely=frame_height)

frame_height += 0.135

check_button_2_var = customtkinter.StringVar(value="off")
check_button_2 = customtkinter.CTkCheckBox(master=app, text="VREES_Fahrplandaten_Strom.xlsx", variable=check_button_2_var, onvalue="on", offvalue="off")
check_button_2.place(relx=0.05, rely=frame_height)

frame_height += 0.135

check_button_3_var = customtkinter.StringVar(value="off")
check_button_3 = customtkinter.CTkCheckBox(master=app, text="VES_SLP_Strom.xlsx", variable=check_button_3_var, onvalue="on", offvalue="off")
check_button_3.place(relx=0.05, rely=frame_height)

frame_height += 0.135

check_button_4_var = customtkinter.StringVar(value="off")
check_button_4 = customtkinter.CTkCheckBox(master=app, text="VES_APM_RLM_SLP.xlsx", variable=check_button_4_var, onvalue="on", offvalue="off")
check_button_4.place(relx=0.05, rely=frame_height)

frame_height += 0.135

# check_button_5_var = customtkinter.StringVar(value="off")
# check_button_5 = customtkinter.CTkCheckBox(master=app, text="Gas_Alloc_ANB.xlsx", variable=check_button_5_var, onvalue="on", offvalue="off")
# check_button_5.place(relx=0.05, rely=frame_height)

# frame_height += 0.1

check_button_6_var = customtkinter.StringVar(value="off")
check_button_6 = customtkinter.CTkCheckBox(master=app, text="Allokation_ANB_Tag_VES_VREES.xlsx", variable=check_button_6_var, onvalue="on", offvalue="off")
check_button_6.place(relx=0.05, rely=frame_height)

ausführen_button = customtkinter.CTkButton(app, text="Die markierten Dateien aktualisieren", command=ausführen, width=120,
                               height=25,
                               corner_radius=8,
                               text_color="white",
                               text_color_disabled="black")
ausführen_button.place(relx=0.2, rely=0.85)

app.mainloop() # halt das Fenster offen
