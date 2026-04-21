import customtkinter as ctk
from tkinter import filedialog, messagebox
import openpyxl
from openpyxl.styles import PatternFill
import pdfplumber
import re
import os
import threading

# Ρυθμίσεις Εμφάνισης
ctk.set_appearance_mode("System")  # Ακολουθεί το θέμα των Windows (Dark/Light)
ctk.set_default_color_theme("blue")

class ModernDataMatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Αντιστοίχιση Δεδομένων Excel & PDF")
        self.root.geometry("650x700")

        self.excel_path = ctk.StringVar()
        self.pdf_path = ctk.StringVar()
        self.export_matches_var = ctk.BooleanVar(value=True)
        self.highlight_excel_var = ctk.BooleanVar(value=True)
        self.search_code_var = ctk.StringVar()

        self.cached_excel_codes = []
        self.cached_pdf_codes = []
        self.last_loaded_excel = ""
        self.last_loaded_pdf = ""

        self.setup_ui()

    def setup_ui(self):
        # Κεντρικό Frame για ωραία περιθώρια
        main_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # --- ΠΛΑΙΣΙΟ 1: Επιλογή Αρχείων ---
        file_frame = ctk.CTkFrame(main_frame)
        file_frame.pack(fill="x", pady=(0, 20))
        ctk.CTkLabel(file_frame, text="1. Επιλογή Αρχείων", font=("Arial", 16, "bold")).pack(anchor="w", padx=15, pady=(10, 5))

        excel_row = ctk.CTkFrame(file_frame, fg_color="transparent")
        excel_row.pack(fill="x", padx=15, pady=5)
        ctk.CTkLabel(excel_row, text="Αρχείο Excel:").pack(side="left", padx=(0, 10))
        ctk.CTkEntry(excel_row, textvariable=self.excel_path, width=320, state='readonly').pack(side="left", padx=(0, 10))
        ctk.CTkButton(excel_row, text="Αναζήτηση", command=self.select_excel, width=100).pack(side="left")

        pdf_row = ctk.CTkFrame(file_frame, fg_color="transparent")
        pdf_row.pack(fill="x", padx=15, pady=(5, 15))
        ctk.CTkLabel(pdf_row, text="Αρχείο PDF:  ").pack(side="left", padx=(0, 10))
        ctk.CTkEntry(pdf_row, textvariable=self.pdf_path, width=320, state='readonly').pack(side="left", padx=(0, 10))
        ctk.CTkButton(pdf_row, text="Αναζήτηση", command=self.select_pdf, width=100).pack(side="left")

        # --- ΠΛΑΙΣΙΟ 2: Μαζικός Έλεγχος & Επιλογές ---
        options_frame = ctk.CTkFrame(main_frame)
        options_frame.pack(fill="x", pady=(0, 20))
        ctk.CTkLabel(options_frame, text="2. Μαζικός Έλεγχος", font=("Arial", 16, "bold")).pack(anchor="w", padx=15, pady=(10, 5))

        ctk.CTkCheckBox(options_frame, text="Δημιουργία Excel ΜΟΝΟ με τα κοινά", variable=self.export_matches_var).pack(anchor="w", padx=15, pady=5)
        ctk.CTkCheckBox(options_frame, text="Δημιουργία νέου Excel με πράσινο Highlight στα κοινά", variable=self.highlight_excel_var).pack(anchor="w", padx=15, pady=(5, 15))

        self.run_btn = ctk.CTkButton(options_frame, text="Εκτέλεση Ελέγχου", command=self.start_matching_thread, font=("Arial", 14, "bold"), fg_color="#2FA572", hover_color="#106A43")
        self.run_btn.pack(pady=(0, 10))

        # Progress Bar & Status (Αρχικά κρυμμένα)
        self.progress = ctk.CTkProgressBar(options_frame, mode="indeterminate", width=400)
        self.progress.set(0)
        self.status_label = ctk.CTkLabel(options_frame, text="", text_color="gray")
        
        self.open_file_btn = ctk.CTkButton(options_frame, text="Άνοιγμα Αρχείου Κοινών", command=self.open_matches_file, fg_color="#1f538d")

        # --- ΠΛΑΙΣΙΟ 3: Μεμονωμένη Αναζήτηση ---
        search_frame = ctk.CTkFrame(main_frame)
        search_frame.pack(fill="x")
        ctk.CTkLabel(search_frame, text="3. Μεμονωμένη Αναζήτηση Εργαζόμενου", font=("Arial", 16, "bold")).pack(anchor="w", padx=15, pady=(10, 5))

        s_row = ctk.CTkFrame(search_frame, fg_color="transparent")
        s_row.pack(fill="x", padx=15, pady=(5, 10))
        ctk.CTkLabel(s_row, text="Εισάγετε Κωδικό:").pack(side="left", padx=(0, 10))
        ctk.CTkEntry(s_row, textvariable=self.search_code_var, width=200).pack(side="left", padx=(0, 10))
        ctk.CTkButton(s_row, text="Αναζήτηση", command=self.start_search_thread, width=100).pack(side="left")
        
        self.search_result_label = ctk.CTkLabel(search_frame, text="", font=("Arial", 14, "bold"))
        self.search_result_label.pack(pady=(0, 15))

    def select_excel(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filepath: self.excel_path.set(filepath)

    def select_pdf(self):
        filepath = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if filepath: self.pdf_path.set(filepath)

    def load_data_to_cache(self):
        excel_file = self.excel_path.get()
        pdf_file = self.pdf_path.get()

        if not excel_file or not pdf_file:
            raise Exception("Παρακαλώ επιλέξτε και τα δύο αρχεία (Excel και PDF) πρώτα.")

        if excel_file == self.last_loaded_excel and pdf_file == self.last_loaded_pdf:
            return

        # Ενημέρωση χρήστη
        self.root.after(0, lambda: self.status_label.configure(text="Διαβάζεται το Excel..."))
        
        self.cached_excel_codes = []
        wb = openpyxl.load_workbook(excel_file, data_only=True)
        sheet = wb.active
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0]:
                self.cached_excel_codes.append(str(row[0]).strip())
        
        # Ενημέρωση χρήστη
        self.root.after(0, lambda: self.status_label.configure(text="Εξαγωγή κειμένου από το PDF... Αυτό μπορεί να πάρει λίγο χρόνο."))
        
        pdf_text = ""
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text: pdf_text += text + "\n"

        regex_pattern = r'\b[A-Z0-9]{5,10}\b'
        self.cached_pdf_codes = re.findall(regex_pattern, pdf_text)

        self.last_loaded_excel = excel_file
        self.last_loaded_pdf = pdf_file

    def start_loading_ui(self):
        self.run_btn.configure(state="disabled")
        self.open_file_btn.pack_forget()
        self.progress.pack(pady=(10, 0))
        self.progress.start()
        self.status_label.pack(pady=(5, 10))

    def stop_loading_ui(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.status_label.pack_forget()
        self.run_btn.configure(state="normal")

    def start_matching_thread(self):
        self.start_loading_ui()
        # Τρέχουμε τη βαριά δουλειά σε άλλο thread για να μην κολλήσει το UI
        threading.Thread(target=self.run_matching_logic, daemon=True).start()

    def run_matching_logic(self):
        try:
            self.load_data_to_cache()
            
            self.root.after(0, lambda: self.status_label.configure(text="Επεξεργασία δεδομένων και δημιουργία αρχείων..."))
            matches = [code for code in self.cached_excel_codes if code in self.cached_pdf_codes]

            if self.export_matches_var.get():
                wb_matches = openpyxl.Workbook()
                ws_matches = wb_matches.active
                ws_matches.title = "Matches"
                ws_matches.append(["Κωδικός Εργαζόμενου", "Κατάσταση"])
                for match in matches: ws_matches.append([match, "Match"])
                
                self.matches_filename = "Matches_Only.xlsx"
                wb_matches.save(self.matches_filename)
                self.root.after(0, lambda: self.open_file_btn.pack(pady=10))

            if self.highlight_excel_var.get():
                excel_file = self.excel_path.get()
                wb_original = openpyxl.load_workbook(excel_file)
                ws_original = wb_original.active
                green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                
                for row in ws_original.iter_rows(min_row=2):
                    cell = row[0]
                    if cell.value and str(cell.value).strip() in self.cached_pdf_codes:
                        cell.fill = green_fill
                
                wb_original.save("Highlighted_Excel.xlsx")

            self.root.after(0, lambda: messagebox.showinfo("Ολοκληρώθηκε", "Η αντιστοίχιση ολοκληρώθηκε επιτυχώς!"))

        except Exception as e:
            self.root.after(0, lambda e=e: messagebox.showerror("Σφάλμα", str(e)))
        finally:
            self.root.after(0, self.stop_loading_ui)

    def open_matches_file(self):
        try:
            os.startfile(self.matches_filename)
        except Exception as e:
            messagebox.showerror("Σφάλμα", f"Δεν ήταν δυνατό το άνοιγμα:\n{str(e)}")

    def start_search_thread(self):
        code_to_search = self.search_code_var.get().strip()
        if not code_to_search:
            self.search_result_label.configure(text="Παρακαλώ γράψτε κωδικό.", text_color="red")
            return

        self.start_loading_ui()
        threading.Thread(target=self.run_search_logic, args=(code_to_search,), daemon=True).start()

    def run_search_logic(self, code_to_search):
        try:
            self.load_data_to_cache()
            
            in_excel = code_to_search in self.cached_excel_codes
            in_pdf = code_to_search in self.cached_pdf_codes

            if in_excel and in_pdf:
                self.root.after(0, lambda: self.search_result_label.configure(text="Ο εργαζόμενος υπάρχει και στα 2 αρχεία", text_color="green"))
            else:
                self.root.after(0, lambda: self.search_result_label.configure(text="Δεν υπάρχει και στα 2 αρχεία ο εργαζόμενος", text_color="red"))
                
        except Exception as e:
            self.root.after(0, lambda e=e: messagebox.showerror("Σφάλμα", str(e)))
        finally:
            self.root.after(0, self.stop_loading_ui)

if __name__ == "__main__":
    app_root = ctk.CTk()
    app = ModernDataMatcherApp(app_root)
    app_root.mainloop()