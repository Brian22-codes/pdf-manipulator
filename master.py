#Tkinter LAnding page, master code 
import tkinter as tk
from tkinter import messagebox
import sys
import subprocess

# Create a lis to store the subprocesses
subprocesses = []

def pdf2word_clicked():
    script_path = "./PDF2WORD.py"
    process = subprocess.Popen(["python", script_path])
    subprocesses.append(process)

def audiopdf_clicked():
    script_path = "./AUDIOPDF.py"
    process = subprocess.Popen(["python", script_path])
    subprocesses.append(process)

def exit_clicked():
    if messagebox.askokcancel("Exit", "Are you sure you want to exit?"):
        # Terminate all subprocesses before exiting
        for process in subprocesses:
            process.terminate()
        sys.exit()

# Create the main window
window = tk.Tk()
window.title("Office Utility Tool")
window.configure(bg="white")

# Create the software namae frame
software_name_frame = tk.Frame(window, bg="white", bd=2, relief=tk.SOLID)
software_name_frame.pack(pady=10)
#Load  ownership note
print("""
    ##############################
    == Intelectual propertyy of ==
    ==      Brian Kiplangat     ==
    ##############################
    """)

software_name_label = tk.Label(software_name_frame, text="OFFICE UTILITY TOOL", font=("Arial", 16, "bold"), bg="white")
software_name_label.pack(padx=20, pady=10)

# Create the buttons
#Creat the PDF TO WORD FUNCTION CALL BUTTON
pdf2word_button = tk.Button(window, text="PDF TO WORD", font=("Arial", 14, "bold"), bg="navy", fg="white", width=20, height=5, command=pdf2word_clicked)
pdf2word_button.pack(pady=10)

#CONVERTER FUNCTION CALL BUTTOn
pdf2word_description = tk.Label(window, text="Convert PDF files to Word (.docx)", font=("Arial", 12), bg="white")
pdf2word_description.pack()

#Audio pdf file triger button
audiopdf_button = tk.Button(window, text="AUDIO PDF", font=("Arial", 14, "bold"), bg="navy", fg="white", width=20, height=5, command=audiopdf_clicked)
audiopdf_button.pack(pady=10)
#Player buttin
audiopdf_description = tk.Label(window, text="Play audio of your PDF files\nSave Audio (mp3)", font=("Arial", 12), bg="white")
audiopdf_description.pack()

# Create the exit button
exit_button = tk.Button(window, text="Exit", font=("Arial", 12), bg="red", fg="white", width=10, command=exit_clicked)
exit_button.pack(anchor=tk.NE, padx=10, pady=10)

# Start the GUI event loop
window.mainloop()
