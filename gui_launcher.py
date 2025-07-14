import tkinter as tk
from tkinter import ttk, messagebox
from bot_runner import run_whatsapp_bot
import threading
import csv
from collections import defaultdict
import re
from PIL import Image, ImageTk

class WhatsAppBotGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("WhatsApp Message Bot - Powered by Barrana")
        self.root.geometry("800x900")
        self.root.resizable(False, False)

        # Company Logo
        self.load_logo()

        self.status_var = tk.StringVar()
        self.status_var.set("Loading sheet and tab info...")

        self.selected_sheet = tk.StringVar()
        self.display_name_map = {}

        # Sheet dropdown (left-aligned)
        tk.Label(root, text="📄 Select Sheet:", font=("Arial", 12, "bold")).pack(pady=(10, 5), padx=15, anchor='w')
        self.sheet_menu = ttk.Combobox(root, values=[], textvariable=self.selected_sheet, width=60)
        self.sheet_menu.bind("<<ComboboxSelected>>", self.update_tabs)
        self.sheet_menu.pack(padx=15, anchor='w')

        # Scrollable checklist frame (left-aligned)
        tk.Label(root, text="🗂️ Select Tabs:", font=("Arial", 12, "bold")).pack(pady=(15, 5), padx=15, anchor='w')
        self.canvas = tk.Canvas(root, height=250)
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.tabs_frame = tk.Frame(self.canvas)

        self.tabs_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.tabs_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True, padx=(15, 0), anchor='w')
        self.scrollbar.pack(side="right", fill="y", padx=(0, 10))

        self.tab_vars = {}

        # AI Prompt Text Area
        prompt_label = ttk.Label(root, text="🤖 AI Message Prompt:", font=("Arial", 12, "bold"))
        prompt_label.pack(padx=10, pady=(10, 5), anchor='w')

        self.prompt_text = tk.Text(root, height=15, width=90)
        self.prompt_text.pack(padx=10, pady=5)
        self.prompt_text.insert("1.0", """
        ---
        Greetings. My name is Omar Bayat, and I am Property Consultant at White & Co., one of the leading British-owned brokerages in Dubai.

        I am reaching out regarding {unit_info}. I currently have a qualified client searching specifically in the building, and wanted to ask if your apartment is available for rent.

        Just last week, I closed over AED 420,000 in rental deals, and as a Super Agent on Property Finder and TruBroker on Bayut, I can give your unit maximum exposure and help secure a reliable tenant quickly.

        If it is already occupied, please feel free to save my details for future opportunities. I would be happy to assist when the time is right.

        Looking forward to hearing from you.

        Best regards,

        Omar Bayat
        White & Co. Real Estate

        Slightly vary the wording professionally for each message. Output ONLY the final message.
""")

        # Run button
        self.run_button = tk.Button(root, text="▶ Run WhatsApp Bot", command=self.run_bot_thread, font=("Arial", 13))
        self.run_button.pack(pady=15)

        # Status label and log
        tk.Label(root, text="📋 Status Log:", font=("Arial", 12, "bold")).pack()
        self.status_text = tk.Text(root, height=15, width=100, bg="#f9f9f9", wrap="word")
        self.status_text.pack(padx=10, pady=5)
        self.setup_text_tags()

        # Exit button
        self.exit_button = tk.Button(root, text="❌ Exit", command=root.quit, font=("Arial", 11))
        self.exit_button.pack(pady=10)

        # Load CSV after GUI widgets are initialized
        self.sheets_and_tabs = self.load_sheets_and_tabs_from_csv()
        self.sheet_menu['values'] = list(self.sheets_and_tabs.keys())

        self.log_to_gui("✅ GUI loaded successfully.")
        self.status_var.set("Select a sheet and tabs to begin.")

    def load_logo(self):
        try:
            logo_image = Image.open("colored-barrana.png")
            logo_image = logo_image.resize((150, 50), Image.LANCZOS)
            logo_photo = ImageTk.PhotoImage(logo_image)
            logo_label = ttk.Label(self.root, image=logo_photo)
            logo_label.image = logo_photo
            logo_label.pack(pady=(10, 10), padx=15, anchor='w')

            # Adding horizontal double line after the logo
            ttk.Separator(self.root, orient='horizontal').pack(fill='x', pady=(0, 3), padx=10)
            ttk.Separator(self.root, orient='horizontal').pack(fill='x', pady=(0, 15), padx=10)

        except Exception as e:
            print("Error loading logo:", e)

    def setup_text_tags(self):
        self.status_text.tag_config("success", foreground="green")
        self.status_text.tag_config("error", foreground="red")
        self.status_text.tag_config("info", foreground="blue")
        self.status_text.tag_config("default", foreground="black")

    def load_sheets_and_tabs_from_csv(self):
        self.log_to_gui("🔄 Loading sheet_tabs_headers_latest.csv...", "info")
        sheets_and_tabs = defaultdict(set)

        try:
            with open("sheet_tabs_headers_latest.csv", mode='r', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                headers = next(reader)

                sheet_idx, tab_idx = None, None
                for i, h in enumerate(headers):
                    h_clean = h.strip().lower()
                    if h_clean in ["sheet", "sheet name"]:
                        sheet_idx = i
                    elif h_clean in ["tab", "tab name"]:
                        tab_idx = i

                if sheet_idx is None or tab_idx is None:
                    raise ValueError("CSV must have 'Sheet' and 'Tab' columns.")

                for row in reader:
                    if len(row) > max(sheet_idx, tab_idx):
                        original_sheet = row[sheet_idx].strip()
                        tab = row[tab_idx].strip()
                        if original_sheet and tab:
                            match = re.match(r"(.*)_([0-9]+)$", original_sheet)
                            if match:
                                sheet_base_name, sheet_number = match.groups()
                                clean_sheet = f"{sheet_number}. {sheet_base_name.replace('_', ' ').title()}"
                            else:
                                clean_sheet = original_sheet.replace('_', ' ').title()

                            sheets_and_tabs[clean_sheet].add(tab)
                            self.display_name_map[clean_sheet] = original_sheet

            self.log_to_gui(f"✅ Loaded {len(sheets_and_tabs)} sheets from CSV.", "success")
        except Exception as e:
            self.log_to_gui(f"❌ Error loading CSV: {e}", "error")
            messagebox.showerror("Error", f"Failed to load sheet_tabs_headers_latest.csv:\n{e}")

        return {k: sorted(v) for k, v in sheets_and_tabs.items()}

    def update_tabs(self, event=None):
        for widget in self.tabs_frame.winfo_children():
            widget.destroy()
        self.tab_vars.clear()

        sheet = self.selected_sheet.get()
        if sheet and sheet in self.sheets_and_tabs:
            for tab in self.sheets_and_tabs[sheet]:
                var = tk.BooleanVar()
                chk = tk.Checkbutton(self.tabs_frame, text=tab, variable=var)
                chk.pack(anchor='w')
                self.tab_vars[tab] = var
            self.log_to_gui(f"📋 Loaded {len(self.tab_vars)} tabs for '{sheet}'", "info")

    def run_bot_thread(self):
        selected_tabs = [tab for tab, var in self.tab_vars.items() if var.get()]
        selected_display_sheet = self.selected_sheet.get()
        original_sheet_name = self.display_name_map.get(selected_display_sheet, selected_display_sheet)

        if not original_sheet_name or not selected_tabs:
            messagebox.showwarning("Selection Required", "Please select a sheet and at least one tab.")
            return

        custom_prompt = self.prompt_text.get("1.0", "end").strip()

        self.log_to_gui(f"🚀 Running bot for Sheet: {selected_display_sheet}, Tabs: {', '.join(selected_tabs)}", "info")
        self.run_button.config(state=tk.DISABLED)

        thread = threading.Thread(target=self.run_bot, args=(original_sheet_name, selected_tabs, custom_prompt))
        thread.start()

    def run_bot(self, sheet_name, selected_tabs, prompt):
        try:
            run_whatsapp_bot(sheet_name, selected_tabs, prompt, log_fn=self.log_to_gui)
        except Exception as e:
            self.log_to_gui(f"❌ Error running bot: {e}", "error")
        finally:
            self.run_button.config(state=tk.NORMAL)
            self.log_to_gui("🏁 Bot finished running.\n", "success")

    def log_to_gui(self, msg, tag="default"):
        self.status_text.insert(tk.END, msg + "\n", tag)
        self.status_text.see(tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppBotGUI(root)
    root.mainloop()


'''WhatsApp Message Bot GUI Launcher


import tkinter as tk
from tkinter import ttk, messagebox
from bot_runner import run_whatsapp_bot
import threading
import csv
from collections import defaultdict
import re

class WhatsAppBotGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("WhatsApp Message Bot")
        self.root.geometry("700x800")
        self.root.resizable(False, False)

        self.status_var = tk.StringVar()
        self.status_var.set("Loading sheet and tab info...")

        self.selected_sheet = tk.StringVar()
        self.display_name_map = {}

        # Sheet dropdown
        tk.Label(root, text="📄 Select Sheet:", font=("Arial", 12, "bold")).pack(pady=(10, 5))
        self.sheet_menu = ttk.Combobox(root, values=[], textvariable=self.selected_sheet, width=60)
        self.sheet_menu.bind("<<ComboboxSelected>>", self.update_tabs)
        self.sheet_menu.pack()

        # Scrollable checklist frame
        tk.Label(root, text="🗂️ Select Tabs:", font=("Arial", 12, "bold")).pack(pady=(15, 5))
        self.canvas = tk.Canvas(root, height=250)
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.tabs_frame = tk.Frame(self.canvas)

        self.tabs_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.tabs_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True, padx=(10, 0))
        self.scrollbar.pack(side="right", fill="y", padx=(0, 10))

        self.tab_vars = {}

        # Run button
        self.run_button = tk.Button(root, text="▶ Run WhatsApp Bot", command=self.run_bot_thread, font=("Arial", 13))
        self.run_button.pack(pady=15)

        # Status label and log
        tk.Label(root, text="📋 Status Log:", font=("Arial", 12, "bold")).pack()
        self.status_text = tk.Text(root, height=20, width=90, bg="#f9f9f9", wrap="word")
        self.status_text.pack(padx=10, pady=5)
        self.setup_text_tags()

        # Exit button
        self.exit_button = tk.Button(root, text="❌ Exit", command=root.quit, font=("Arial", 11))
        self.exit_button.pack(pady=10)

        # Load CSV after GUI widgets are initialized
        self.sheets_and_tabs = self.load_sheets_and_tabs_from_csv()
        self.sheet_menu['values'] = list(self.sheets_and_tabs.keys())

        self.log_to_gui("✅ GUI loaded successfully.")
        self.status_var.set("Select a sheet and tabs to begin.")

    def setup_text_tags(self):
        self.status_text.tag_config("success", foreground="green")
        self.status_text.tag_config("error", foreground="red")
        self.status_text.tag_config("info", foreground="blue")
        self.status_text.tag_config("default", foreground="black")

    def load_sheets_and_tabs_from_csv(self):
        self.log_to_gui("🔄 Loading sheet_tabs_headers_latest.csv...", "info")
        sheets_and_tabs = defaultdict(set)

        try:
            with open("sheet_tabs_headers_latest.csv", mode='r', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                headers = next(reader)

                sheet_idx, tab_idx = None, None
                for i, h in enumerate(headers):
                    h_clean = h.strip().lower()
                    if h_clean in ["sheet", "sheet name"]:
                        sheet_idx = i
                    elif h_clean in ["tab", "tab name"]:
                        tab_idx = i

                if sheet_idx is None or tab_idx is None:
                    raise ValueError("CSV must have 'Sheet' and 'Tab' columns.")

                for row in reader:
                    if len(row) > max(sheet_idx, tab_idx):
                        original_sheet = row[sheet_idx].strip()
                        tab = row[tab_idx].strip()
                        if original_sheet and tab:
                            match = re.match(r"(.*)_([0-9]+)$", original_sheet)
                            if match:
                                sheet_base_name, sheet_number = match.groups()
                                clean_sheet = f"{sheet_number}. {sheet_base_name.replace('_', ' ').title()}"
                            else:
                                clean_sheet = original_sheet.replace('_', ' ').title()

                            sheets_and_tabs[clean_sheet].add(tab)
                            self.display_name_map[clean_sheet] = original_sheet  # Use self here

            self.log_to_gui(f"✅ Loaded {len(sheets_and_tabs)} sheets from CSV.", "success")
        except Exception as e:
            self.log_to_gui(f"❌ Error loading CSV: {e}", "error")
            messagebox.showerror("Error", f"Failed to load sheet_tabs_headers_latest.csv:\n{e}")
    
        return {k: sorted(v) for k, v in sheets_and_tabs.items()}

    def update_tabs(self, event=None):
        for widget in self.tabs_frame.winfo_children():
            widget.destroy()
        self.tab_vars.clear()

        sheet = self.selected_sheet.get()
        if sheet and sheet in self.sheets_and_tabs:
            for tab in self.sheets_and_tabs[sheet]:
                var = tk.BooleanVar()
                chk = tk.Checkbutton(self.tabs_frame, text=tab, variable=var)
                chk.pack(anchor='w')
                self.tab_vars[tab] = var
            self.log_to_gui(f"📋 Loaded {len(self.tab_vars)} tabs for '{sheet}'", "info")

    def run_bot_thread(self):
        selected_tabs = [tab for tab, var in self.tab_vars.items() if var.get()]
        selected_display_sheet = self.selected_sheet.get()
        original_sheet_name = self.display_name_map.get(selected_display_sheet, selected_display_sheet)

        if not original_sheet_name or not selected_tabs:
            messagebox.showwarning("Selection Required", "Please select a sheet and at least one tab.")
            return

        self.log_to_gui(f"🚀 Running bot for Sheet: {selected_display_sheet}, Tabs: {', '.join(selected_tabs)}", "info")
        self.run_button.config(state=tk.DISABLED)

        thread = threading.Thread(target=self.run_bot, args=(original_sheet_name, selected_tabs))
        thread.start()

    def run_bot(self, sheet_name, selected_tabs):
        try:
            run_whatsapp_bot(selected_sheet_name=sheet_name, selected_tabs=selected_tabs, log_fn=self.log_to_gui)
        except Exception as e:
            self.log_to_gui(f"❌ Error running bot: {e}", "error")
        finally:
            self.run_button.config(state=tk.NORMAL)
            self.log_to_gui("🏁 Bot finished running.\n", "success")

    def log_to_gui(self, msg, tag="default"):
        self.status_text.insert(tk.END, msg + "\n", tag)
        self.status_text.see(tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppBotGUI(root)
    root.mainloop()
'''