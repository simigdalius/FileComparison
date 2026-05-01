import customtkinter as ctk
from tkinter import filedialog, messagebox
import openpyxl
from openpyxl.styles import PatternFill
import pdfplumber
import re
import os
import threading
import time

# Ρυθμίσεις Εμφάνισης
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class ModernDataMatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Αντιστοίχιση Δεδομένων Excel & PDF")
        # Μεγαλώσαμε το παράθυρο σε πλάτος (1100) για να χωράει η προεπισκόπηση δεξιά
        self.root.geometry("1100x750")

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
        # container
        main_container = ctk.CTkFrame(self.root, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=20, pady=20)

        # aristerh sthlh
        left_frame = ctk.CTkFrame(main_container, fg_color="transparent")
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

        # file selection
        file_frame = ctk.CTkFrame(left_frame)
        file_frame.pack(fill="x", pady=(0, 20))
        ctk.CTkLabel(file_frame, text="1. Επιλογή Αρχείων", font=("Arial", 16, "bold")).pack(anchor="w", padx=15, pady=(10, 5))

        excel_row = ctk.CTkFrame(file_frame, fg_color="transparent")
        excel_row.pack(fill="x", padx=15, pady=5)
        ctk.CTkLabel(excel_row, text="Αρχείο Excel:").pack(side="left", padx=(0, 10))
        ctk.CTkEntry(excel_row, textvariable=self.excel_path, width=280, state='readonly').pack(side="left", padx=(0, 10))
        ctk.CTkButton(excel_row, text="Αναζήτηση", command=self.select_excel, width=90).pack(side="left")

        pdf_row = ctk.CTkFrame(file_frame, fg_color="transparent")
        pdf_row.pack(fill="x", padx=15, pady=(5, 15))
        ctk.CTkLabel(pdf_row, text="Αρχείο PDF:  ").pack(side="left", padx=(0, 10))
        ctk.CTkEntry(pdf_row, textvariable=self.pdf_path, width=280, state='readonly').pack(side="left", padx=(0, 10))
        ctk.CTkButton(pdf_row, text="Αναζήτηση", command=self.select_pdf, width=90).pack(side="left")

        # mazikos elegxos
        options_frame = ctk.CTkFrame(left_frame)
        options_frame.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(options_frame, text="2. Μαζικός Έλεγχος", font=("Arial", 16, "bold")).pack(anchor="w", padx=15, pady=(10, 5))

        ctk.CTkCheckBox(options_frame, text="Δημιουργία Excel ΜΟΝΟ με τα κοινά", variable=self.export_matches_var).pack(anchor="w", padx=15, pady=5)
        ctk.CTkCheckBox(options_frame, text="Πράσινο Highlight στα κοινά (Στο αρχικό Excel)", variable=self.highlight_excel_var).pack(anchor="w", padx=15, pady=(5, 15))

        self.run_btn = ctk.CTkButton(options_frame, text="Εκτέλεση Ελέγχου", command=self.start_matching_thread, font=("Arial", 14, "bold"), fg_color="#2FA572", hover_color="#106A43")
        self.run_btn.pack(pady=(0, 10))

        self.loading_frame = ctk.CTkFrame(options_frame, height=60, fg_color="transparent")
        self.loading_frame.pack(fill="x", pady=5)
        self.loading_frame.pack_propagate(False) 

        self.progress = ctk.CTkProgressBar(self.loading_frame, mode="indeterminate", width=380)
        self.progress.set(0)
        self.status_label = ctk.CTkLabel(self.loading_frame, text="", text_color="gray")
        
        self.open_file_btn = ctk.CTkButton(options_frame, text="Άνοιγμα Αρχείου Κοινών", command=self.open_matches_file, fg_color="#1f538d")

        # memonomenh anazhthsh
        search_frame = ctk.CTkFrame(left_frame)
        search_frame.pack(fill="x", pady=(30, 0))
        ctk.CTkLabel(search_frame, text="3. Μεμονωμένη Αναζήτηση Εργαζόμενου", font=("Arial", 16, "bold")).pack(anchor="w", padx=15, pady=(10, 5))

        s_row = ctk.CTkFrame(search_frame, fg_color="transparent")
        s_row.pack(fill="x", padx=15, pady=(5, 10))
        ctk.CTkLabel(s_row, text="Εισάγετε Κωδικό:").pack(side="left", padx=(0, 10))
        ctk.CTkEntry(s_row, textvariable=self.search_code_var, width=180).pack(side="left", padx=(0, 10))
        ctk.CTkButton(s_row, text="Αναζήτηση", command=self.start_search_thread, width=90).pack(side="left")
        
        self.search_result_label = ctk.CTkLabel(search_frame, text="", font=("Arial", 14, "bold"))
        self.search_result_label.pack(pady=(0, 15))


        right_frame = ctk.CTkFrame(main_container)
        right_frame.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        ctk.CTkLabel(right_frame, text="Προεπισκόπηση Κοινών (Matches)", font=("Arial", 16, "bold")).pack(pady=(15, 10))
        
        # Πεδίο κειμένου που λειτουργεί σαν "πίνακας"
        self.preview_box = ctk.CTkTextbox(right_frame, wrap="none", font=("Consolas", 14))
        self.preview_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.preview_box.insert("0.0", "Κάντε εκτέλεση ελέγχου για να\nεμφανιστούν τα αποτελέσματα εδώ...")
        self.preview_box.configure(state="disabled") # Το κλειδώνουμε για να μην γράφει ο χρήστης

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

        self.root.after(0, lambda: self.status_label.configure(text="Διαβάζεται το Excel..."))
        
        self.cached_excel_codes = []
        wb = openpyxl.load_workbook(excel_file, data_only=True)
        sheet = wb.active
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0]:
                self.cached_excel_codes.append(str(row[0]).strip())
        
        pdf_text = ""
        with pdfplumber.open(pdf_file) as pdf:
            total_pages = len(pdf.pages)
            for i, page in enumerate(pdf.pages):
                self.root.after(0, lambda p=i+1, t=total_pages: self.status_label.configure(text=f"Διάβασμα PDF: Σελίδα {p} από {t}..."))
                text = page.extract_text()
                if text: 
                    pdf_text += text + "\n"
                time.sleep(0.05)

        self.root.after(0, lambda: self.status_label.configure(text="Αναζήτηση κωδικών στο κείμενο..."))
        #megethos kwdikoy pou psaxnw
        regex_pattern = r'\b[A-Z]+-\d+-\d+-\d+\b'
        self.cached_pdf_codes = re.findall(regex_pattern, pdf_text)

        self.last_loaded_excel = excel_file
        self.last_loaded_pdf = pdf_file

    def start_loading_ui(self):
        self.run_btn.configure(state="disabled")
        self.open_file_btn.pack_forget()
        self.progress.pack(pady=(5, 0))
        self.progress.start()
        self.status_label.pack(pady=(2, 0))
        
        # Καθαρισμός προεπισκόπησης κατά την εκκίνηση
        self.preview_box.configure(state="normal")
        self.preview_box.delete("0.0", "end")
        self.preview_box.insert("end", "Γίνεται επεξεργασία, παρακαλώ περιμένετε...")
        self.preview_box.configure(state="disabled")

    def stop_loading_ui(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.status_label.pack_forget()
        self.run_btn.configure(state="normal")

    def update_preview_ui(self, matches):
        #enhmerwsh pinka sta dejia
        self.preview_box.configure(state="normal")
        self.preview_box.delete("0.0", "end")

        self.preview_box.insert("end", f"{'Κωδικός Εργαζόμενου':<22} | {'Κατάσταση'}\n")
        self.preview_box.insert("end", "-" * 40 + "\n")
        
        if not matches:
            self.preview_box.insert("end", "Δεν βρέθηκαν κοινοί κωδικοί.\n")
        else:
            for match in matches:
                self.preview_box.insert("end", f"{match:<22} | Match\n")
                
        self.preview_box.configure(state="disabled")

    def start_matching_thread(self):
        self.start_loading_ui()
        threading.Thread(target=self.run_matching_logic, daemon=True).start()

    def run_matching_logic(self):
        try:
            self.load_data_to_cache()
            
            self.root.after(0, lambda: self.status_label.configure(text="Επεξεργασία δεδομένων..."))
            matches = [code for code in self.cached_excel_codes if code in self.cached_pdf_codes]

            self.root.after(0, lambda m=matches: self.update_preview_ui(m))

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
                
                wb_original.save(excel_file)

            self.root.after(0, lambda: messagebox.showinfo("Ολοκληρώθηκε", f"Η αντιστοίχιση ολοκληρώθηκε επιτυχώς!\nΒρέθηκαν {len(matches)} κοινοί κωδικοί."))

        except Exception as e:
            error_msg = str(e)
            if "Permission denied" in error_msg:
                error_msg = "Δεν ήταν δυνατή η αποθήκευση. Βεβαιωθείτε ότι το αρχείο Excel είναι ΚΛΕΙΣΤΟ και ξαναδοκιμάστε."
            self.root.after(0, lambda e=error_msg: messagebox.showerror("Σφάλμα", e))
            
            # periptwsh p xtyphsei lathos
            self.root.after(0, lambda: self.update_preview_ui([]))
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
            self.root.after(0, lambda: self.update_preview_ui([]))

if __name__ == "__main__":
    app_root = ctk.CTk()
    app = ModernDataMatcherApp(app_root)
    app_root.mainloop()