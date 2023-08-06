import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
from ttkbootstrap import Style
from openpyxl import Workbook, load_workbook
import os
import random

DATA_FILE = "quiz_data.xlsx"

class QuizApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Kanji Quiz")
        self.quiz_categories = []
        self.selected_categories = []
        self.current_category = ""
        self.current_quiz = None
        self.score = 0
        self.total_questions = 0
        self.shown_quizzes = {}
        self.quiz_data = {}
        self.quiz_direction = 'kanji to meaning' 

        self.category_var = tk.StringVar(value=self.selected_categories)
        self.category_listbox = tk.Listbox(root, listvariable=self.category_var, selectmode=tk.MULTIPLE, font=('Arial', 12))

        # Adding Scrollbar
        self.category_scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.category_listbox.yview)
        self.category_listbox.configure(yscrollcommand=self.category_scrollbar.set)

        self.switch_mode_button = ttk.Button(root, text='Cambia modalità', style='warning.TButton', command=self.switch_mode)
        
        self.question_label = ttk.Label(root, font=('Arial', 24))

        self.answer_entry = ttk.Entry(root, font=('Arial', 18))

        self.submit_button = ttk.Button(root, text='Invia', style='danger.TButton', command=self.check_answer, state=tk.DISABLED)

        self.next_button = ttk.Button(root, text='Prossima domanda', style='success.TButton', command=self.next_question, state=tk.DISABLED)

        self.score_label = ttk.Label(root, font=('Arial', 16))

        self.add_category_button = ttk.Button(root, text='Aggiungi Categoria', style='info.TButton', command=self.open_add_category_window)

        self.edit_category_button = ttk.Button(root, text='Modifica Categoria', style='info.TButton', command=self.open_edit_category_window)

        self.add_quiz_button = ttk.Button(root, text='  Aggiungi Quiz  ', style='info.TButton', command=self.open_add_quiz_window)

        self.edit_quiz_button = ttk.Button(root, text='  Modifica Quiz  ', style='info.TButton', command=self.open_edit_quiz_window)

        self.category_listbox.bind('<<ListboxSelect>>', self.load_quiz)

        self.load_quiz_data()

        self.create_layout()

    def create_layout(self):
        self.add_category_button.grid(row=0, column=0, pady=10, padx=10)
        self.edit_category_button.grid(row=1, column=0, pady=10, padx=10)
        self.category_listbox.grid(row=0, column=1, rowspan=4, pady=10, padx=10, sticky='nsew')
        self.category_scrollbar.grid(row=0, column=2, rowspan=4, pady=10, sticky='ns')
        self.add_quiz_button.grid(row=0, column=3, pady=10, padx=10)
        self.edit_quiz_button.grid(row=1, column=3, pady=10, padx=10)

        self.switch_mode_button.grid(row=4, column=0, columnspan=4, pady=10)
        self.question_label.grid(row=5, column=0, columnspan=4, pady=10)
        self.answer_entry.grid(row=6, column=0, columnspan=4, pady=10)
        self.submit_button.grid(row=7, column=0, columnspan=4)
        self.next_button.grid(row=8, column=0, columnspan=4, pady=10)
        self.score_label.grid(row=9, column=0, columnspan=4)


    def open_add_category_window(self):
        category = simpledialog.askstring("Aggiungi Categoria", "Inserisci il nome della categoria:")
        if category:
            if category not in self.quiz_categories:
                self.quiz_categories.append(category)
                self.category_listbox.insert(tk.END, category)
                self.quiz_data[category] = []
                messagebox.showinfo('Categoria Aggiunta', 'La categoria è stata aggiunta con successo!')
                self.save_quiz_data()
            else:
                messagebox.showerror('Errore', 'La categoria esiste già.')


    def open_edit_category_window(self):
        selected_index = self.category_listbox.curselection()
        if not selected_index:
            messagebox.showerror('Errore', 'Seleziona una categoria da modificare.')
            return

        category_index = selected_index[0]
        old_category_name = self.quiz_categories[category_index]
        new_category_name = simpledialog.askstring("Modifica Categoria", f"Modifica il nome della categoria '{old_category_name}':")
        if new_category_name:
            if new_category_name not in self.quiz_categories:
                self.quiz_categories[category_index] = new_category_name
                self.category_listbox.delete(category_index)
                self.category_listbox.insert(category_index, new_category_name)
                self.quiz_data[new_category_name] = self.quiz_data.pop(old_category_name)  # Rinomina la chiave nel dizionario
                self.save_quiz_data()
            else:
                messagebox.showerror('Errore', 'La categoria esiste già.')


    def open_edit_quiz_window(self):
        selected_index = self.category_listbox.curselection()
        if not selected_index:
            messagebox.showerror('Errore', 'Seleziona una categoria per modificare un quiz.')
            return

        category_index = selected_index[0]
        selected_category = self.quiz_categories[category_index]
        if selected_category not in self.quiz_data:
            messagebox.showinfo('Nessun quiz', 'La categoria selezionata non contiene ancora quiz. Aggiungine uno prima di modificarlo.')
            return

        edit_quiz_window = tk.Toplevel(self.root)
        edit_quiz_window.title('Modifica Quiz')

        quiz_listbox = tk.Listbox(edit_quiz_window, font=('Arial', 12), width=50)
        quiz_listbox.pack(pady=10)

        for quiz in self.quiz_data[selected_category]:
            quiz_listbox.insert(tk.END, f"{quiz['kanji']} | {quiz['meaning']}")

        def update_quiz():
            selected_quiz_index = quiz_listbox.curselection()
            if not selected_quiz_index:
                messagebox.showerror('Errore', 'Seleziona un quiz da modificare.')
                return

            quiz_index = selected_quiz_index[0]
            old_quiz = self.quiz_data[selected_category][quiz_index]

            new_kanji = kanji_entry.get()
            new_meaning = meaning_entry.get()
            new_romaji = romaji_entry.get()
            if new_kanji and new_meaning and new_romaji:
                self.quiz_data[selected_category][quiz_index] = {'kanji': new_kanji, 'romaji': new_romaji, 'meaning': new_meaning}
                self.save_quiz_data()
                edit_quiz_window.destroy()
            else:
                messagebox.showerror('Errore', 'Inserisci il kanji, il romaji e il significato.')


        kanji_label = tk.Label(edit_quiz_window, text='Kanji:', font=('Arial', 14))
        kanji_label.pack()

        kanji_entry = tk.Entry(edit_quiz_window, font=('Arial', 14))
        kanji_entry.pack()

        romaji_label = tk.Label(edit_quiz_window, text='Romaji:', font=('Arial', 14))
        romaji_label.pack(pady=10)

        romaji_entry = tk.Entry(edit_quiz_window, font=('Arial', 14))
        romaji_entry.pack(pady=10)
        
        meaning_label = tk.Label(edit_quiz_window, text='Significato:', font=('Arial', 14))
        meaning_label.pack()

        meaning_entry = tk.Entry(edit_quiz_window, font=('Arial', 14))
        meaning_entry.pack()

        edit_button = tk.Button(edit_quiz_window, text='Modifica', font=('Arial', 14), command=update_quiz)
        edit_button.pack(pady=10)



    def load_quiz_data(self):
        try:
            if not os.path.isfile(DATA_FILE):
                wb = Workbook()
                ws = wb.active
                ws.title = 'Generale'
                wb.save(DATA_FILE)
                self.quiz_data['Generale'] = []
            else:
                wb = load_workbook(DATA_FILE)
                for sheet in wb.sheetnames:
                    if sheet not in self.quiz_data:
                        self.quiz_data[sheet] = []
                    if sheet != 'Generale' and sheet not in self.quiz_categories:
                        self.quiz_categories.append(sheet)
                    for row in wb[sheet].values:
                        kanji, romaji, meaning, category = (row + (None, None, None, None))[:4]
                        if romaji or meaning:
                            if not category:
                                category = "Generale"
                            self.quiz_data[sheet].append({'kanji': kanji, 'romaji': romaji, 'meaning': meaning, 'category': category})
            self.category_var.set(self.quiz_categories)
        except Exception as e:
            messagebox.showerror('Errore', f"Errore durante il caricamento dei dati del quiz: {str(e)}")




    def save_quiz_data(self):
        try:
            # Controlla se il file esiste già
            if os.path.isfile(DATA_FILE):
                # Se esiste, carica il foglio di lavoro
                wb = load_workbook(DATA_FILE)
            else:
                # Se non esiste, crea un nuovo foglio di lavoro
                wb = Workbook()

            for category, quizzes in self.quiz_data.items():
                if category in wb.sheetnames:
                    # se il foglio esiste già nel foglio di lavoro, usalo
                    ws = wb[category]
                else:
                    # se il foglio non esiste nel foglio di lavoro, crealo
                    ws = wb.create_sheet(title=category)

                # Cancella le righe esistenti nel foglio di lavoro
                for row in ws.iter_rows(min_row=ws.min_row, max_col=ws.max_column, max_row=ws.max_row):
                    for cell in row:
                        cell.value = None

                # Aggiungi nuove righe
                for quiz in quizzes:
                    ws.append([quiz['kanji'], quiz['romaji'], quiz['meaning'], quiz['category']])

            wb.save(DATA_FILE)

        except PermissionError:
            messagebox.showerror("Errore di salvataggio", "Impossibile salvare i dati del quiz. "
                                                          "Assicurati che il file non sia aperto in un altro programma e riprova.")
        except Exception as e:
            messagebox.showerror("Errore di salvataggio", f"Si è verificato un errore durante il salvataggio dei dati del quiz: {str(e)}")


    def switch_mode(self):
        if self.quiz_direction == 'kanji to meaning':
            self.quiz_direction = 'meaning to kanji'
            messagebox.showinfo('Modalità cambiata', 'Ora devi indovinare il kanji dal significato.')
        else:
            self.quiz_direction = 'kanji to meaning'
            messagebox.showinfo('Modalità cambiata', 'Ora devi indovinare il significato dal kanji.')
        self.next_question()  # Mostra un nuovo quiz dopo il cambio di modalita

    def next_random_quiz(self):
        selected_categories = self.category_listbox.curselection()
        if not selected_categories:
            messagebox.showerror('Errore', 'Seleziona una o più categorie.')
            return

        self.current_category = self.quiz_categories[random.choice(selected_categories)]
    
        if self.current_category not in self.quiz_data:
            self.quiz_data[self.current_category] = []

        if self.current_category not in self.shown_quizzes:
            self.shown_quizzes[self.current_category] = set()

        if len(self.shown_quizzes[self.current_category]) == len(self.quiz_data[self.current_category]):
            # If we have shown all quizzes in this category, reset the set
            self.shown_quizzes[self.current_category] = set()

        remaining_quizzes = [q for q in self.quiz_data[self.current_category] if str(q) not in self.shown_quizzes[self.current_category]]

        if remaining_quizzes:
            quiz = random.choice(remaining_quizzes)
            self.shown_quizzes[self.current_category].add(str(quiz))
            return quiz
        else:
            messagebox.showinfo('No Quizzes', 'La categoria selezionata non contiene ancora quiz.')
            return None


    def check_answer(self):
        user_answer = self.answer_entry.get()
        if not user_answer:
            messagebox.showerror('Errore', 'Inserisci una risposta.')
            return

        if self.quiz_direction == 'meaning to kanji':
            correct_answer = self.current_quiz['kanji'] if self.current_quiz['kanji'] else self.current_quiz['romaji']
        else:
            correct_answer = self.current_quiz['meaning']

        if user_answer.lower() == correct_answer.lower():
            self.score += 1
        else:
            messagebox.showinfo('Risposta Sbagliata', f"La risposta corretta era: {correct_answer}")

        self.total_questions += 1
        self.score_label['text'] = f"Punteggio: {self.score}/{self.total_questions}"
        self.answer_entry.delete(0, tk.END)
        self.next_button['state'] = tk.NORMAL
        self.submit_button['state'] = tk.DISABLED


    def load_quiz(self, event):
        selected_categories = [self.quiz_categories[i] for i in self.category_listbox.curselection()]
        if selected_categories:
            self.selected_categories = selected_categories
            self.next_question()

    def next_question(self):
        next_quiz = self.next_random_quiz()
        if next_quiz is not None:
            self.current_quiz = next_quiz
            if self.quiz_direction == 'kanji to meaning':
                if next_quiz['kanji']:
                    self.question_label['text'] = f"Quale è il significato di questo kanji: {next_quiz['kanji']} ({next_quiz['romaji']})?"
                else:
                    self.question_label['text'] = f"Quale è il significato di questo romaji: {next_quiz['romaji']}?"
            else:
                if next_quiz['kanji']:
                    self.question_label['text'] = f"Quale kanji rappresenta questo significato: {next_quiz['meaning']}?"
                else:
                    self.question_label['text'] = f"Quale romaji rappresenta questo significato: {next_quiz['meaning']}?"

            self.next_button['state'] = tk.DISABLED
            self.submit_button['state'] = tk.NORMAL




    def open_add_quiz_window(self):
        selected_index = self.category_listbox.curselection()
        if not selected_index:
            messagebox.showerror('Errore', 'Seleziona una categoria per aggiungere un quiz.')
            return

        category_index = selected_index[0]
        selected_category = self.quiz_categories[category_index]

        add_quiz_window = tk.Toplevel(self.root)
        add_quiz_window.title('Aggiungi Quiz')

        kanji_label = tk.Label(add_quiz_window, text='Kanji:', font=('Arial', 14))
        kanji_label.pack()

        kanji_entry = tk.Entry(add_quiz_window, font=('Arial', 14))
        kanji_entry.pack()

        romaji_label = tk.Label(add_quiz_window, text='Romaji:', font=('Arial', 14))
        romaji_label.pack()

        romaji_entry = tk.Entry(add_quiz_window, font=('Arial', 14))
        romaji_entry.pack()

        meaning_label = tk.Label(add_quiz_window, text='Significato:', font=('Arial', 14))
        meaning_label.pack()

        meaning_entry = tk.Entry(add_quiz_window, font=('Arial', 14))
        meaning_entry.pack()

        def add_quiz():
            kanji = kanji_entry.get()
            meaning = meaning_entry.get()
            romaji = romaji_entry.get()
            if romaji or meaning:
                self.quiz_data[selected_category].append({'kanji': kanji, 'romaji': romaji, 'meaning': meaning, 'category': selected_category})
                self.save_quiz_data()
                add_quiz_window.destroy()
            else:
                messagebox.showerror('Errore', 'Inserisci sia il kanji che il significato.')

        add_button = tk.Button(add_quiz_window, text='Aggiungi', font=('Arial', 14), command=add_quiz)
        add_button.pack(pady=10)


if __name__ == "__main__":
    root = tk.Tk()
    style = Style(theme="cyborg")
    QuizApp(root)
    root.mainloop()

