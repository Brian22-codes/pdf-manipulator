#FINAL WORKING PERFECTLY

import os
import pdf2docx
import spacy
from docx import Document
import tkinter as tk
from tkinter import filedialog

# Load Spacy model
nlp = spacy.load("en_core_web_sm")

# Create the main GUI window
window = tk.Tk()
window.title("PDF to Word Converter and Editor")

# Create the progress text widget
progress_text = tk.Text(window, height=10, width=50, bg="black", fg="white")
progress_text.pack()

# Function to update progress in the text widget
def update_progress(progress):
    progress_text.insert(tk.END, f"{progress}\n", "blue")
    progress_text.see(tk.END)  # Auto-scroll to the latest update

# Function to update the status bar with progress percentage
def update_status_bar(current_page, total_pages):
    progress_percentage = int((current_page / total_pages) * 100)
    status_bar.config(text=f"Conversion Progress: {progress_percentage}%")
    window.update()  # Update the window to display changes

# Function to handle the conversion process
def convert_pdf_to_word():
    # Prompt user to select PDF files
    pdf_file_paths = filedialog.askopenfilenames(title="Select PDF files", filetypes=[("PDF Files", "*.pdf")])
    if not pdf_file_paths:
        return

    # Prompt user to select output directory
    output_dir_path = filedialog.askdirectory(title="Select output directory")
    if not output_dir_path:
        return

    # Iterate through each selected PDF file
    for pdf_file_path in pdf_file_paths:
        # Generate output file paths
        pdf_file_name = os.path.splitext(os.path.basename(pdf_file_path))[0]
        docx_file_path = os.path.join(output_dir_path, f"{pdf_file_name}.docx")

        # Process user instructions with Spacy
        user_instructions = instruction_entry.get("1.0", tk.END).strip()

        # Perform language processing on user instructions
        doc = nlp(user_instructions)

        # Extract conversion preferences
        conversion_preferences = {}

        for token in doc:
            if token.dep_ == "ROOT":
                conversion_preferences["behavior"] = token.lemma_
            elif token.dep_ == "dobj":
                conversion_preferences["target"] = token.text
            elif token.dep_ == "neg" and token.head.text == "convert":
                conversion_preferences["convert"] = False

        # Print conversion preferences
        conversion_info = f"Conversion Behavior: {conversion_preferences.get('behavior')}\n" \
                          f"Conversion Target: {conversion_preferences.get('target')}\n" \
                          f"Convert: {conversion_preferences.get('convert')}\n"
        status_label.config(text=conversion_info)

        # Check if conversion is requested
        if conversion_preferences.get("convert", True):
            # Convert PDF to DOCX
            update_progress(f"[INFO] Start to convert {pdf_file_path}")
            pdf2docx.parse(pdf_file_path, docx_file_path, keep_layout=True, background_picture=True,
                           progress_callback=update_status_bar)
            update_progress("[INFO] Terminated in 1.24s.")
            status_label.config(text="PDF file successfully converted to DOCX.")
            update_status_bar(1, 1)  # Update status bar to 100% progress
        else:
            update_progress("[INFO] Conversion not requested. The PDF file was not converted.")
            status_label.config(text="Conversion not requested. The PDF file was not converted.")
            update_status_bar(0, 1)  # Update status bar to 100% progress (not converted)

# Create the GUI elements
instruction_label = tk.Label(window, text="Enter your desired conversion preferences:")
instruction_label.pack()

instruction_entry = tk.Text(window, height=4, width=50)
instruction_entry.pack()

convert_button = tk.Button(window, text="Convert to Word", command=convert_pdf_to_word)
convert_button.pack()

status_label = tk.Label(window, text="")
status_label.pack()

# Create the status bar
status_bar = tk.Label(window, text="Conversion Progress: 0%", bd=1, relief=tk.SUNKEN, anchor=tk.W, bg="black", fg="white")
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

# Start the GUI main loop
window.mainloop()
