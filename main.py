import difflib
import pandas as pd
import geopandas as gpd
import numpy as np
from tkinter import Tk, filedialog, Button, Label, Entry, StringVar, messagebox, ttk, DoubleVar
import threading
import sys
import re

def choose_unique_representation(strings):
    return max(strings, key=len)

def find_similar_words(input_string, string_list, threshold):
    similar_words = [input_string]  # Aggiungiamo la stringa di input alla lista
    for word in string_list:
        similarity_ratio = difflib.SequenceMatcher(None, input_string, word).ratio()
        if similarity_ratio >= threshold:
            similar_words.append(word)
    return similar_words

def compute_similarity(args):
    stringa1, stringa2 = args
    matcher = difflib.SequenceMatcher(None, stringa1, stringa2)
    similarita = matcher.ratio()
    return similarita

def create_similarity_matrix(lista1, lista2, progress_label, progress_var):
    num_stringhe1 = len(lista1)
    num_stringhe2 = len(lista2)
    current_iteration = 0
    total_iterations = num_stringhe1 * num_stringhe2

    # Inizializza la matrice di similarità con zeri
    matrice_similarita = np.zeros((num_stringhe1, num_stringhe2))

    # Calcola la matrice di similarità su un singolo thread
    for i in range(num_stringhe1):
        for j in range(num_stringhe2):
            matrice_similarita[i, j] = compute_similarity((lista1[i], lista2[j]))
            current_iteration += 1
            progress = (current_iteration / total_iterations) * 100
            progress_var.set(progress)
            root.update()  # Aggiornamento della GUI

    return matrice_similarita

def to_lower(input_list):
    input_list = [string for string in input_list if string is not None]
    input_list = [string.lower() for string in input_list]
    
    return input_list

def remove_unuseful_entries(input_list):
    return [string for string in input_list if string is not None and (len(string) >= 5 and string != "" and string != " " and string != "(vuoto)")]

def remove_unuseful_characters(input_list):
    result = []
    pattern = re.compile(r'[^a-zA-Z]')  # Questo pattern corrisponde a qualsiasi cosa che non sia una lettera

    for string in input_list:
        # Rimuovi spazi bianchi, numeri e caratteri speciali
        clean_string = re.sub(pattern, '', string)
        result.append(clean_string)

    return result

def preprocess(input_list):
    #tmp_list = remove_unuseful_entries(input_list)
    tmp = to_lower(input_list)
    return remove_unuseful_characters(tmp)

def browse_file(entry_var):
    Tk().withdraw()  # non vogliamo una GUI completa, quindi manteniamo la finestra principale nascosta
    filename = filedialog.askopenfilename(filetypes=[("File Excel", "*.xlsx")])
    entry_var.set(filename)

def browse_shapefile(entry_var):
    Tk().withdraw()
    filename = filedialog.askopenfilename(filetypes=[("Shapefile", "*.shx")])
    entry_var.set(filename)

def process_data(file_path, shapefile_path, root, progress_label, progress_var):
    file = pd.read_excel(file_path)

    shp = gpd.read_file(shapefile_path)

    lista_strade_da_bonficare = file['Etichette di riga']
    lista_strade_osm = shp['name']

    list1 = lista_strade_osm.tolist()
    list2 = lista_strade_da_bonficare.tolist()
    
    progress_label.config(text="Preprocessing lista strade osm...")
    list1_p = preprocess(list1)
    progress_var.set(0)
    progress_label.config(text="Preprocessing lista strade excel...")
    list2_p = preprocess(list2)
    progress_var.set(0)

    progress_label.config(text="Creazione matrice di similarità...")
    matrice_similarita = create_similarity_matrix(list1_p, list2_p, progress_label, progress_var)
    progress_var.set(0)

    list1_index = [string for string in list1 if string is not None]

    df = pd.DataFrame(matrice_similarita, index=list1_index, columns=list2)
    #df.to_excel("similarity_matrix.xlsx")

    # Trova l'indice del valore massimo solo se è maggiore di 0.85
    max_indices = df.apply(lambda col: col.idxmax() if col.max() > 0.85 else "", axis=0)

    # Aggiungi i valori massimi alla colonna "JOIN" nel file iniziale
    file['JOIN'] = max_indices.values

    # Lista delle stringhe che non hanno valore nella colonna "JOIN"
    empty_join_strings = np.array(file[file['JOIN'] == ""]['Etichette di riga'])

    progress_label.config(text="Postprocessing...")
    total_iterations = len(empty_join_strings)
    current_iteration = 0
    
    for string in empty_join_strings:
        if string != "":
            # Trova tutte le stringhe simili alla stringa corrente
            similar_strings = find_similar_words(string, empty_join_strings, 0.85)

            if len(similar_strings) > 1:
                # Trova una rappresentazione unica per il gruppo di stringhe simili
                unique_representation = choose_unique_representation(similar_strings)

                # Assegna la rappresentazione unica a tutte le stringhe simili
                file.loc[file['Etichette di riga'].isin(similar_strings), 'MOST SIGNIFICANT'] = unique_representation

            current_iteration += 1
            progress = (current_iteration / total_iterations) * 100
            progress_var.set(progress)
            root.update()  # Aggiornamento della GUI

    progress_var.set(85)

    progress_label.config(text="Finalizzazione...")
    # Salva il file aggiornato
    file.to_excel("TARI_con_JOIN.xlsx")
    progress_var.set(100)

    # Visualizza il messaggio di successo
    messagebox.showinfo("Successo", "Operazione di bonfica andata a buon fine!")

    # Resetta la barra di avanzamento
    progress_var.set(0)

if __name__ == "__main__":
    root = Tk()
    root.title("Seleziona File")
    root.geometry("500x300")  # Aggiornato per dare spazio alla label sopra la progress bar

    file_path_var = StringVar()
    shapefile_path_var = StringVar()

    file_label = Label(root, text="Seleziona File Excel:")
    file_label.grid(row=0, column=0, pady=10, padx=10, sticky="w")

    file_entry = Entry(root, textvariable=file_path_var, state="readonly", width=40)
    file_entry.grid(row=0, column=1, pady=10, padx=10)

    file_button = Button(root, text="Sfoglia", command=lambda: browse_file(file_path_var))
    file_button.grid(row=0, column=2, pady=10, padx=10)

    shapefile_label = Label(root, text="Seleziona Shapefile (.shx):")
    shapefile_label.grid(row=1, column=0, pady=10, padx=10, sticky="w")

    shapefile_entry = Entry(root, textvariable=shapefile_path_var, state="readonly", width=40)
    shapefile_entry.grid(row=1, column=1, pady=10, padx=10)

    shapefile_button = Button(root, text="Sfoglia", command=lambda: browse_shapefile(shapefile_path_var))
    shapefile_button.grid(row=1, column=2, pady=10, padx=10)

    progress_var = DoubleVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, length=400, mode="determinate")
    progress_bar.grid(row=2, column=0, columnspan=3, pady=20, padx=10)

    progress_label = Label(root, text="")
    progress_label.grid(row=3, column=0, columnspan=3, pady=10, padx=10)  # Posizionato sopra la progress bar

    ok_button = Button(root, text="OK", command=lambda: threading.Thread(target=process_data, args=(file_path_var.get(), shapefile_path_var.get(), root, progress_label, progress_var)).start())
    ok_button.grid(row=4, column=1, pady=20)

    cancel_button = Button(root, text="Annulla", command=lambda: sys.exit(0))  # Chiudi il programma quando si clicca su "Annulla"
    cancel_button.grid(row=4, column=2, pady=20)

    root.protocol("WM_DELETE_WINDOW", lambda: sys.exit(0))  # Chiudi il programma quando si clicca sulla "X"
    root.mainloop()
