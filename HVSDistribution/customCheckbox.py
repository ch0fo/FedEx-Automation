import tkinter as tk
from tkinter import messagebox

class CustomDialog(tk.Toplevel):
    def __init__(self, parent, title=None):
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()

        if title:
            self.title(title)
        
        self.parent = parent
        self.result = None

        self.body_frame = tk.Frame(self)
        self.body_frame.pack(padx=5, pady=5)

        self.checkbox_value = tk.BooleanVar()

        self.body(self.body_frame)
        self.buttonbox()
    
    def body(self, master):
        tk.Label(master, text="Please connect to Zscaler VPN, or disconnect from Checkpoint Mobile VPN, and reconnect once the upload has started.").pack()
        tk.Checkbutton(master, text="Do not show this message again", variable=self.checkbox_value).pack()
    
    def buttonbox(self):
        box = tk.Frame(self)
        
        tk.Button(box, text="Upload", width=10, command=self.ok).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(box, text="Cancel", width=10, command=self.cancel).pack(side=tk.LEFT, padx=5, pady=5)

        box.pack()

    def ok(self):
        self.result = True
        self.withdraw()
        self.update_idletasks()
        self.parent.focus_set()
        self.destroy()

    def cancel(self):
        self.result = False
        self.withdraw()
        self.update_idletasks()
        self.parent.focus_set()
        self.destroy()

    def get_checkbox_value(self):
        return self.checkbox_value.get()

def show_custom_warning(root):
    dialog = CustomDialog(root, title="Warning")
    root.wait_window(dialog)
    
    return dialog.result, dialog.get_checkbox_value()

