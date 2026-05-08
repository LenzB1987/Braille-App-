import customtkinter as ctk
from tkinter import filedialog, messagebox
import os, sys, json

# Attempt to import docx for Word support
try:
    import docx
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

class UEBTranslator:
    """A Pure-Python Unified English Braille Translator (Grade 1 & 2)"""
    def __init__(self):
        self.braille_map = {
            'a': '⠁', 'b': '⠃', 'c': '⠉', 'd': '⠙', 'e': '⠑', 'f': '⠋', 'g': '⠛', 'h': '⠓', 'i': '⠊', 'j': '⠚',
            'k': '⠅', 'l': '⠇', 'm': '⠍', 'n': '⠝', 'o': '⠕', 'p': '⠏', 'q': '⠟', 'r': '⠗', 's': '⠎', 't': '⠞',
            'u': '⠥', 'v': '⠧', 'w': '⠺', 'x': '⠭', 'y': '⠽', 'z': '⠵',
            '1': '⠁', '2': '⠃', '3': '⠉', '4': '⠙', '5': '⠑', '6': '⠋', '7': '⠛', '8': '⠓', '9': '⠊', '0': '⠚',
            ' ': '⠀', '.': '⠲', ',': '⠂', ';': '⠆', ':': '⠒', '!': '⠖', '?': '⠦', '"': '⠐⠂', "'": '⠄', '-': '⠤',
            '(': '⠐⠣', ')': '⠐⠜', '/': '⠸⠌', '\\': '⠸⠡', '@': '⠈⠁', '#': '⠼', '$': '⠈⠎', '%': '⠨⠴', '&': '⠈⠿', '*': '⠐⠔'
        }
        
        # Grade 2 Whole Word Contractions
        self.contractions_whole = {
            "but": "⠃", "can": "⠉", "do": "⠙", "every": "⠑", "from": "⠋", "go": "⠛", "have": "⠓", "just": "⠚",
            "knowledge": "⠅", "like": "⠇", "more": "⠍", "not": "⠝", "people": "⠏", "quite": "⠟", "rather": "⠗",
            "so": "⠎", "that": "⠞", "us": "⠥", "very": "⠧", "will": "⠺", "it": "⠭", "you": "⠽", "as": "⠵",
            "and": "⠯", "for": "⠿", "of": "⠷", "the": "⠮", "with": "⠾", "was": "⠴", "were": "⠶", "his": "⠒", "in": "⠊"
        }
        
        # Grade 2 Part Word Contractions (order matters)
        self.contractions_part = [
            ("and", "⠯"), ("for", "⠿"), ("of", "⠷"), ("the", "⠮"), ("with", "⠾"),
            ("ch", "⠡"), ("sh", "⠱"), ("th", "⠹"), ("wh", "⠱"), ("ou", "⠳"),
            ("st", "⠌"), ("ar", "⠜"), ("ed", "⠫"), ("er", "⠻"), ("gh", "⠣"),
            ("ow", "⠪"), ("ing", "⠬")
        ]

    def translate(self, text, grade="g2"):
        if not text: return ""
        
        lines = text.split('\n')
        translated_lines = []
        
        for line in lines:
            words = line.split(' ')
            translated_words = []
            
            for word_full in words:
                if not word_full:
                    translated_words.append("")
                    continue
                
                # Separate punctuation
                word = "".join(filter(str.isalnum, word_full)).lower()
                prefix = ""
                suffix = ""
                
                # Find punctuation at end
                punc_idx = len(word_full)
                while punc_idx > 0 and not word_full[punc_idx-1].isalnum():
                    punc_idx -= 1
                suffix = word_full[punc_idx:]
                
                # Find punctuation at start
                start_idx = 0
                while start_idx < len(word_full) and not word_full[start_idx].isalnum():
                    start_idx += 1
                prefix = word_full[:start_idx]
                word = word_full[start_idx:punc_idx]

                res = ""
                
                # Grade 2 Logic
                if grade == "g2" and word.lower() in self.contractions_whole:
                    res = self.contractions_whole[word.lower()]
                else:
                    processed = word.lower()
                    if grade == "g2":
                        for pattern, replacement in self.contractions_part:
                            processed = processed.replace(pattern, replacement)
                    
                    is_num = False
                    for i, char in enumerate(processed):
                        orig_char = word[i] if i < len(word) else char
                        
                        if char.isdigit():
                            if not is_num:
                                res += '⠼'
                                is_num = True
                            res += self.braille_map.get(char, char)
                        else:
                            is_num = False
                            if orig_char.isupper():
                                res += '⠠'
                            
                            if ord(char) >= 0x2800: # Already braille
                                res += char
                            else:
                                res += self.braille_map.get(char.lower(), char)
                
                # Add punctuation
                final_word = ""
                for c in prefix: final_word += self.braille_map.get(c, c)
                final_word += res
                for c in suffix: final_word += self.braille_map.get(c, c)
                
                translated_words.append(final_word)
            
            translated_lines.append("⠀".join(translated_words))
            
        return "\n".join(translated_lines)

class BrailleStudioPro(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configuration
        self.title("Lenz Improved Braille Version")
        self.geometry("1200x850")
        ctk.set_appearance_mode("dark")
        
        self.translator = UEBTranslator()
        self.debounce_id = None
        
        # Layout
        self.grid_columnconfigure((0, 1), weight=1)
        self.grid_rowconfigure(1, weight=1)

        # --- Header ---
        self.header = ctk.CTkFrame(self, height=70, corner_radius=0, fg_color="#1a1a1a")
        self.header.grid(row=0, column=0, columnspan=2, sticky="ew")
        
        self.logo = ctk.CTkLabel(self.header, text="⠿ Lenz Improved Braille App", font=("Inter", 24, "bold"), text_color="#3b82f6")
        self.logo.pack(side="left", padx=30)
        
        # Right Side Buttons
        self.btn_print = ctk.CTkButton(self.header, text="Print", width=80, fg_color="#2563eb", hover_color="#1d4ed8", command=self.print_dialog)
        self.btn_print.pack(side="right", padx=(10, 30))
        
        self.btn_save = ctk.CTkButton(self.header, text="Save .brf", width=80, command=self.save_file)
        self.btn_save.pack(side="right", padx=10)
        
        self.btn_import = ctk.CTkButton(self.header, text="Import File", width=100, command=self.import_file)
        self.btn_import.pack(side="right", padx=10)

        # --- Toolbar ---
        self.toolbar = ctk.CTkFrame(self, height=50, corner_radius=0, fg_color="#242424")
        self.toolbar.grid(row=2, column=0, columnspan=2, sticky="ew")
        
        # Grade Toggle
        self.grade_var = ctk.StringVar(value="g2")
        self.grade_label = ctk.CTkLabel(self.toolbar, text="Translation:", font=("Inter", 12, "bold"))
        self.grade_label.pack(side="left", padx=(30, 10))
        
        self.grade_menu = ctk.CTkOptionMenu(self.toolbar, values=["Grade 1 (Literal)", "Grade 2 (Contracted)"], 
                                          command=self.change_grade, width=180)
        self.grade_menu.set("Grade 2 (Contracted)")
        self.grade_menu.pack(side="left", padx=5)
        
        # Color Picker
        self.color_label = ctk.CTkLabel(self.toolbar, text="Braille Color:", font=("Inter", 12, "bold"))
        self.color_label.pack(side="left", padx=(30, 10))
        
        self.color_menu = ctk.CTkOptionMenu(self.toolbar, values=["Emerald", "Sky Blue", "White", "Amber", "Ruby"], 
                                           command=self.change_color, width=120)
        self.color_menu.set("Emerald")
        self.color_menu.pack(side="left", padx=5)
        
        # Font Size
        self.size_label = ctk.CTkLabel(self.toolbar, text="Size:", font=("Inter", 12, "bold"))
        self.size_label.pack(side="left", padx=(30, 10))
        self.size_slider = ctk.CTkSlider(self.toolbar, from_=20, to=80, command=self.change_size, width=150)
        self.size_slider.set(42)
        self.size_slider.pack(side="left", padx=5)

        # --- Main Editors ---
        # Input Box
        self.input_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.input_frame.grid(row=1, column=0, sticky="nsew", padx=(20, 10), pady=20)
        
        self.input_label = ctk.CTkLabel(self.input_frame, text="SOURCE TEXT", font=("Inter", 10, "bold"), text_color="#666")
        self.input_label.pack(anchor="w", pady=(0, 5))
        
        self.txt_input = ctk.CTkTextbox(self.input_frame, font=("Inter", 18), border_width=1, border_color="#333")
        self.txt_input.pack(fill="both", expand=True)
        self.txt_input.bind("<KeyRelease>", self.on_key)
        
        # Output Box
        self.output_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.output_frame.grid(row=1, column=1, sticky="nsew", padx=(10, 20), pady=20)
        
        self.output_label = ctk.CTkLabel(self.output_frame, text="BRAILLE OUTPUT", font=("Inter", 10, "bold"), text_color="#666")
        self.output_label.pack(anchor="w", pady=(0, 5))
        
        self.txt_output = ctk.CTkTextbox(self.output_frame, font=("Segoe UI Symbol", 42), state="disabled", text_color="#4ade80")
        self.txt_output.pack(fill="both", expand=True)

        # --- Footer ---
        self.footer = ctk.CTkFrame(self, height=30, corner_radius=0, fg_color="#111")
        self.footer.grid(row=3, column=0, columnspan=2, sticky="ew")
        
        self.stats = ctk.CTkLabel(self.footer, text="Words: 0 | Characters: 0 | Braille Cells: 0", font=("Inter", 11), text_color="#555")
        self.stats.pack(side="left", padx=20)
        
        self.engine_status = ctk.CTkLabel(self.footer, text="● Engine: Pure Python UEB v2.0", font=("Inter", 11), text_color="#4ade80")
        self.engine_status.pack(side="right", padx=20)

    # --- Logic ---

    def on_key(self, event):
        if self.debounce_id:
            self.after_cancel(self.debounce_id)
        self.debounce_id = self.after(150, self.update_translation)

    def update_translation(self):
        content = self.txt_input.get("1.0", "end-1c")
        grade = "g2" if "Grade 2" in self.grade_menu.get() else "g1"
        
        braille = self.translator.translate(content, grade)
        
        self.txt_output.configure(state="normal")
        self.txt_output.delete("1.0", "end")
        self.txt_output.insert("1.0", braille)
        self.txt_output.configure(state="disabled")
        
        # Update Stats
        words = len(content.split())
        chars = len(content)
        cells = len(braille.replace("\n", "").replace(" ", ""))
        self.stats.configure(text=f"Words: {words} | Characters: {chars} | Braille Cells: {cells}")

    def change_grade(self, choice):
        self.update_translation()

    def change_color(self, choice):
        colors = {
            "Emerald": "#4ade80",
            "Sky Blue": "#3b82f6",
            "White": "#ffffff",
            "Amber": "#f59e0b",
            "Ruby": "#ef4444"
        }
        self.txt_output.configure(text_color=colors.get(choice, "#4ade80"))

    def change_size(self, val):
        self.txt_output.configure(font=("Segoe UI Symbol", int(val)))

    def import_file(self):
        file_types = [("Text files", "*.txt")]
        if HAS_DOCX:
            file_types.append(("Word files", "*.docx"))
            
        path = filedialog.askopenfilename(filetypes=file_types)
        if not path: return
        
        try:
            if path.endswith(".docx"):
                doc = docx.Document(path)
                text = "\n".join([p.text for p in doc.paragraphs])
            else:
                with open(path, 'r', encoding='utf-8') as f:
                    text = f.read()
            
            self.txt_input.delete("1.0", "end")
            self.txt_input.insert("1.0", text)
            self.update_translation()
        except Exception as e:
            messagebox.showerror("Import Error", f"Could not read file: {e}")

    def save_file(self):
        path = filedialog.asksaveasfilename(defaultextension=".brf", 
                                          filetypes=[("Braille Ready Format", "*.brf"), ("Text file", "*.txt")])
        if not path: return
        
        try:
            content = self.txt_output.get("1.0", "end-1c")
            with open(path, 'w', encoding='utf-8') as f:
                f.write(content)
            messagebox.showinfo("Success", "Braille document saved successfully.")
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save file: {e}")

    def print_dialog(self):
        # On Windows/Mac, the easiest "print" for a custom GUI is to export to a temporary file 
        # and open it with the system print dialog or just notify the user.
        messagebox.showinfo("Print", "To print your Braille document:\n\n1. Save the file as .brf or .txt\n2. Open with Notepad/TextEdit\n3. Press Ctrl+P / Cmd+P")

if __name__ == "__main__":
    app = BrailleStudioPro()
    app.mainloop()
