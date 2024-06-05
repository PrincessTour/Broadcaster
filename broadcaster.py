import pywhatkit

import datetime

import time

import pandas as pd

import tkinter as tk

from tkinter import ttk

import threading

 

def add_plus_if_missing(string):

    if not string.startswith("+"):

        string = "+" + string

    return string

 

def execute_backend():

    contacts = pd.read_excel("Contacts.xlsx")

    list = contacts.iloc[:, 2]

 

    # parse the setup file

 

    c = 0

    for num in list:

        num = str(num)

        num = add_plus_if_missing(num)

        print(num)

        try:

            if contacts.iloc[c, 3] == 'no': # il contatto non ha ancora ricevuto il messaggio

                if nome_img == None:

                    if "cliente" in testo:

                        new_testo = testo.replace("cliente", contacts.iloc[c,0])

                    else:

                        new_testo = testo

                    # Invio solo messaggio di testo

                    now = datetime.datetime.now()

                    time.sleep(5)

                    #pywhatkit.sendwhatmsg(num, new_testo, now.hour, int(now.minute)+2, 15, True, 10)

                    pywhatkit.sendwhatmsg_instantly(num, new_testo, 15, True, 10)

                else:

                    # Invio immagine + messaggio di testo

                    if "cliente" in testo:

                        new_testo = testo.replace("cliente", contacts.iloc[c,0])

                    else:

                        new_testo = testo

                    pywhatkit.sendwhats_image(num, nome_img, new_testo, 15, True, 16 )

                # il messaggio è stato inviato allo specifico utente, settiamo il flag a 'si'

                contacts.at[c, "Sent"] = 'si'

                contacts.to_excel("Contacts.xlsx", index=False)

                time.sleep(10)

        except Exception as e:

            # qualcosa è andato storto, salviamo sull'excel tutti i contatti a cui siamo riusciti ad inviare so far

            print(e)

            contacts.to_excel("Contacts.xlsx", index=False)

        c += 1

    # abbiamo finito la lista, resettiamo l'excel per un nuovo broadcast

    contacts['Sent'] = 'no'

    contacts.to_excel("Contacts.xlsx", index=False)

 

# Funzione per abilitare/disabilitare il campo del nome immagine

def toggle_image_name_field():

    if image_flag_var.get():

        image_name_entry.config(state='normal')

    else:

        image_name_entry.config(state='disabled')

 

# Funzione per gestire il pulsante di invio

def submit():

    global testo, nome_img

    testo = text_entry.get("1.0", tk.END).strip()  # Ottenere il testo dal widget Text

    nome_img = image_name_entry.get() if image_flag_var.get() else None

    # Stampa per debug

    print(f"Testo: {testo}")

    print(f"Nome Immagine: {nome_img}")

    # Chiamata alla funzione per mostrare il messaggio di invio

    show_broadcast_message()

 

# Funzione per mostrare il messaggio di invio broadcast

def show_broadcast_message():

    # Rimuove tutti i widget attuali dalla finestra

    for widget in root.winfo_children():

        widget.destroy()

    # Mostra il messaggio di invio broadcast in corso

    broadcast_label = ttk.Label(root, text="Invio broadcast in corso...", font=("Helvetica", 16, "bold"))

    broadcast_label.pack(expand=True, padx=20, pady=20)

   

    # Pianifica l'esecuzione del backend dopo 10 millisecondi

    root.after(10, execute_backend)

   

    # Pianifica l'aggiornamento dell'interfaccia utente dopo il completamento del backend

    root.after(1000, update_ui_after_broadcast_sent)

 

def update_ui_after_broadcast_sent():

    # Rimuove tutti i widget attuali dalla finestra dopo il completamento del backend

    for widget in root.winfo_children():

        widget.destroy()

    broadcast_label = ttk.Label(root, text="Invio broadcast completato! Grazie!", font=("Helvetica", 16, "bold"))

    broadcast_label.pack(expand=True, padx=20, pady=20)

 

# Creazione della finestra principale

root = tk.Tk()

root.title("Whatsapp Broadcast")

root.geometry("1700x700")

 

# Stile

style = ttk.Style()

style.configure("TLabel", font=("Helvetica", 12), padding=5)

style.configure("TEntry", padding=5)

style.configure("TButton", font=("Helvetica", 12), padding=5)

style.configure("TCheckbutton", font=("Helvetica", 12), padding=5)

 

# Frame principale

main_frame = ttk.Frame(root, padding="10 20 10 10")  # Ridotto il padding a sinistra

main_frame.pack(expand=True, fill="both")

 

# Aggiunta del manuale utente

manual_label = ttk.Label(main_frame, text="MANUALE UTENTE:\n- fare una prima connessione a whatsapp web col telefono aziendale. Una volta connesso, chiudere whatsapp web\n- il telefono deve essere connesso alla stessa rete internet del computer\n- nel testo, se si scrive la parola 'cliente', nel messaggio finale essa verrà sostituita con il nome della persona a cui si sta inviando il messaggio.\n Esempio: Ciao cliente --> Ciao Edoardo", font=("Helvetica", 12))

manual_label.grid(column=0, row=0, columnspan=2, padx=10, pady=10, sticky='W')

 

# Campo di testo

text_label = ttk.Label(main_frame, text="Testo:")

text_label.grid(column=0, row=1, padx=(0, 20), pady=10, sticky='NW')  # Aumentato il padding a destra

 

text_entry = tk.Text(main_frame, wrap='word', width=50, height=10)

text_entry.grid(column=1, row=1, padx=10, pady=10, sticky='W')

 

# Flag immagine

image_flag_var = tk.BooleanVar()

image_flag_label = ttk.Label(main_frame, text="Allega Immagine:")

image_flag_label.grid(column=0, row=2, padx=(0, 20), pady=10, sticky='W')  # Aumentato il padding a destra

 

image_flag_check = ttk.Checkbutton(main_frame, variable=image_flag_var, command=toggle_image_name_field)

image_flag_check.grid(column=1, row=2, padx=10, pady=10, sticky='W')

 

# Campo nome immagine (disabilitato di default)

image_name_label = ttk.Label(main_frame, text="Path Immagine:")

image_name_label.grid(column=0, row=3, padx=(0, 20), pady=10, sticky='W')  # Aumentato il padding a destra

 

image_name_entry = ttk.Entry(main_frame, width=48, state='disabled')

image_name_entry.grid(column=1, row=3, padx=10, pady=10, sticky='W')

 

# Pulsante di invio

submit_button = ttk.Button(main_frame, text="Invia", command=submit)

submit_button.grid(column=1, row=4, padx=10, pady=20, sticky='E')

 

# Configurazione delle colonne per il ridimensionamento

main_frame.columnconfigure(0, weight=1)

main_frame.columnconfigure(1, weight=3)

 

# Loop principale dell'interfaccia

root.mainloop()

 

# Variabili globali per memorizzare i dati inseriti dall'utente

testo = ""

nome_img = None