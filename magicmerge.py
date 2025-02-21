#!/usr/bin/python3

#Magic Merge by Alayna Ferdarko - Created 21 February, 2025.

import os
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD

def browse_files():
    files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
    add_files(files)

def add_files(files):
    """ Adds files to the listbox and updates size estimate """
    for file in files:
        if file.endswith(".pdf") and file not in pdf_files:
            pdf_files.append(file)
            listbox.insert(tk.END, file)
    update_size_estimate()

def remove_selected():
    selected_indices = listbox.curselection()
    for index in reversed(selected_indices):
        pdf_files.pop(index)
        listbox.delete(index)
    update_size_estimate()

def move_up():
    selected_indices = listbox.curselection()
    for index in selected_indices:
        if index > 0:
            pdf_files[index], pdf_files[index - 1] = pdf_files[index - 1], pdf_files[index]
            refresh_listbox()
            listbox.selection_set(index - 1)
            break

def move_down():
    selected_indices = listbox.curselection()
    for index in selected_indices:
        if index < len(pdf_files) - 1:
            pdf_files[index], pdf_files[index + 1] = pdf_files[index + 1], pdf_files[index]
            refresh_listbox()
            listbox.selection_set(index + 1)
            break

def refresh_listbox():
    listbox.delete(0, tk.END)
    for file in pdf_files:
        listbox.insert(tk.END, file)

def update_size_estimate():
    total_size = sum(os.path.getsize(f) for f in pdf_files)
    size_label.config(text=f"Estimated Size: {total_size / (1024 * 1024):.2f} MB")

def browse_save_location():
    """ Let user choose where to save the merged PDF """
    file_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                             filetypes=[("PDF Files", "*.pdf")],
                                             title="Choose Save Location")
    if file_path:
        save_location_entry.delete(0, tk.END)
        save_location_entry.insert(0, file_path)

def merge_pdfs():
    if not pdf_files:
        messagebox.showerror("Error", "No PDFs selected.")
        return
    
    output_path = save_location_entry.get().strip()
    if not output_path:
        messagebox.showerror("Error", "Select a save location for the merged PDF.")
        return
    
    progress_bar.start()
    
    merger = fitz.open()
    for pdf in pdf_files:
        with fitz.open(pdf) as doc:
            merger.insert_pdf(doc)
    
    merger.save(output_path)
    merger.close()
    
    progress_bar.stop()
    messagebox.showinfo("Success", f"Merged PDF saved as:\n{output_path}")

def clear_list():
    pdf_files.clear()
    listbox.delete(0, tk.END)
    update_size_estimate()


def on_drag(event):
    """ Handles files dropped into the window """
    files = app.tk.splitlist(event.data)  # Get dropped files
    add_files(files)

# Initialize TkinterDnD to enable drag-and-drop
app = TkinterDnD.Tk()
app.title("Magic Merge by Alayna Ferdarko")

pdf_files = []
pdf_viewer_frame = None  # Initialize as None to avoid reference issues

frame = tk.Frame(app)
frame.pack(pady=10)

listbox = tk.Listbox(frame, width=50, height=10)
listbox.pack(side=tk.LEFT, padx=5)
scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL, command=listbox.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
listbox.config(yscrollcommand=scrollbar.set)

# Enable drag-and-drop for the listbox
listbox.drop_target_register(DND_FILES)
listbox.dnd_bind("<<Drop>>", on_drag)

btn_browse = tk.Button(app, text="Browse PDFs", command=browse_files)
btn_browse.pack(pady=5)

btn_remove = tk.Button(app, text="Remove Selected", command=remove_selected)
btn_remove.pack(pady=5)

btn_up = tk.Button(app, text="Move Up", command=move_up)
btn_up.pack(pady=5)

btn_down = tk.Button(app, text="Move Down", command=move_down)
btn_down.pack(pady=5)

size_label = tk.Label(app, text="Estimated Size: 0 MB")
size_label.pack()

# Save Location Entry & Browse Button
save_location_entry = tk.Entry(app, width=50)
save_location_entry.pack(pady=5)

btn_save_location = tk.Button(app, text="Browse Save Location", command=browse_save_location)
btn_save_location.pack(pady=5)

progress_bar = ttk.Progressbar(app, mode="indeterminate")
progress_bar.pack(pady=5)

btn_merge = tk.Button(app, text="Merge PDFs", command=merge_pdfs)
btn_merge.pack(pady=5)

btn_clear = tk.Button(app, text="Clear List", command=clear_list)
btn_clear.pack(pady=5)

app.mainloop()
