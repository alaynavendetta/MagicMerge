#!/usr/bin/python3

'''
Magic Merge by Alayna Ferdarko
Originally created 21 February, 2025.
Updated on 24 March, 2025.

Version 2.1 Overview:
Added ability to split PDFs.
'''

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import fitz  # PyMuPDF

# --- Functions ---
def merge_pdfs():
    input_files = file_list.get(0, tk.END)
    if not input_files:
        messagebox.showerror("Error", "No PDFs selected for merging.")
        return

    output_path = save_location_entry.get().strip()
    if not output_path:
        messagebox.showerror("Error", "Select a save location for the merged PDF.")
        return

    progress.start()

    try:
        merged_pdf = fitz.open()
        for file in input_files:
            with fitz.open(file) as pdf:
                merged_pdf.insert_pdf(pdf)

        merged_pdf.save(output_path)
        merged_pdf.close()

        messagebox.showinfo("Success", f"PDFs merged successfully!\nSaved as: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

    progress.stop()

def browse_files():
    files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
    if files:
        for file in files:
            file_list.insert(tk.END, file)

def remove_selected():
    selected_files = file_list.curselection()
    for index in reversed(selected_files):
        file_list.delete(index)

def browse_save_location():
    save_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                             filetypes=[("PDF Files", "*.pdf")],
                                             title="Save Merged PDF As")
    if save_path:
        save_location_entry.delete(0, tk.END)
        save_location_entry.insert(0, save_path)

def split_pdf():
    input_path = split_file_path.get().strip()
    page_range_str = range_entry.get().strip()
    output_path = split_save_location_entry.get().strip()

    if not input_path or not os.path.isfile(input_path):
        messagebox.showerror("Error", "Select a valid PDF to split.")
        return

    if not page_range_str:
        messagebox.showerror("Error", "Enter a page range.")
        return

    if not output_path:
        messagebox.showerror("Error", "Select a save location for the split PDF.")
        return

    split_progress.start()

    try:
        with fitz.open(input_path) as doc:
            total_pages = doc.page_count

            # Parse only one range for simplicity: start-end
            if '-' not in page_range_str:
                messagebox.showerror("Error", "Enter a range in the format 'start-end'.")
                split_progress.stop()
                return

            start_str, end_str = page_range_str.split('-')
            start_page = int(start_str.strip()) - 1  # Convert to 0-indexed
            end_page = int(end_str.strip()) - 1

            if start_page < 0 or end_page >= total_pages or start_page > end_page:
                split_progress.stop()
                messagebox.showerror("Error", "Page range is out of bounds.")
                return

            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=start_page, to_page=end_page)
            new_doc.save(output_path)
            new_doc.close()

        messagebox.showinfo("Success", f"Split PDF saved as:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

    split_progress.stop()

def browse_split_file():
    file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file:
        split_file_path.set(file)

def browse_split_save_location():
    file_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                             filetypes=[("PDF Files", "*.pdf")],
                                             title="Save Split PDF As")
    if file_path:
        split_save_location_entry.delete(0, tk.END)
        split_save_location_entry.insert(0, file_path)

# --- UI Setup ---
app = tk.Tk()
app.title("Magic Merge v.2 by Alayna Ferdarko")
app.geometry("600x700")

# --- Merge Section ---
merge_frame = tk.LabelFrame(app, text="Merge PDFs", padx=10, pady=10)
merge_frame.pack(pady=10)

file_list = tk.Listbox(merge_frame, width=70, height=8)
file_list.grid(row=0, column=0, columnspan=3, pady=5)

browse_button = tk.Button(merge_frame, text="Add PDFs", command=browse_files)
browse_button.grid(row=1, column=0, pady=5)

remove_button = tk.Button(merge_frame, text="Remove Selected", command=remove_selected)
remove_button.grid(row=1, column=1, pady=5)

save_location_entry = tk.Entry(merge_frame, width=50)
save_location_entry.grid(row=2, column=0, pady=5)

save_browse_button = tk.Button(merge_frame, text="Browse Save Location", command=browse_save_location)
save_browse_button.grid(row=2, column=1, pady=5)

progress = ttk.Progressbar(merge_frame, mode="indeterminate")
progress.grid(row=3, column=0, columnspan=3, pady=5)

merge_button = tk.Button(merge_frame, text="Merge PDFs", command=merge_pdfs)
merge_button.grid(row=4, column=0, columnspan=3, pady=5)

# --- Split Section ---
split_frame = tk.LabelFrame(app, text="Split PDF", padx=10, pady=10)
split_frame.pack(pady=10)

split_file_path = tk.StringVar()

split_file_button = tk.Button(split_frame, text="Select PDF to Split", command=browse_split_file)
split_file_button.grid(row=0, column=0, pady=5)

split_file_label = tk.Label(split_frame, textvariable=split_file_path, wraplength=400)
split_file_label.grid(row=1, column=0, pady=5)

range_label = tk.Label(split_frame, text="Enter Page Range (e.g., 1-3,5,7-9):")
range_label.grid(row=2, column=0, pady=5)

range_entry = tk.Entry(split_frame, width=50)
range_entry.grid(row=3, column=0, pady=5)

split_save_location_entry = tk.Entry(split_frame, width=50)
split_save_location_entry.grid(row=4, column=0, pady=5)

split_save_button = tk.Button(split_frame, text="Browse Save Location", command=browse_split_save_location)
split_save_button.grid(row=5, column=0, pady=5)

split_progress = ttk.Progressbar(split_frame, mode="indeterminate")
split_progress.grid(row=6, column=0, pady=5)

split_button = tk.Button(split_frame, text="Split PDF", command=lambda: split_pdf())
split_button.grid(row=7, column=0, pady=5)

app.mainloop()
