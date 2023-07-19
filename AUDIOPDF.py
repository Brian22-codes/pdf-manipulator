#FINAL WORKING PERFECTLY
#banner
print("""
    ###############################
    == Intelectual propertyy of ==
    ==      Brian Kiplangat     ==
    ###############################
    """)
import PyPDF2
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from threading import Thread
import pyttsx3
import os

class PDFReader:
    def __init__(self):
        self.file_path = ""
        self.text = ""
        self.selected_text = ""

        self.engine = pyttsx3.init()
        self.engine.setProperty('rate', 150)  # Adjust the reading speed as needed

        self.window = tk.Tk()
        self.window.title("PDF Reader")
        self.window.geometry("600x400")

        self.canvas = tk.Canvas(self.window)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar = tk.Scrollbar(self.window, command=self.canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.frame, anchor='nw')

        self.text_display = tk.Text(self.frame, height=15, width=140)
        self.text_display.pack(padx=10, pady=10)

        self.options_frame = tk.Frame(self.frame)
        self.options_frame.pack(padx=10, pady=10)

        self.voice_label = tk.Label(self.options_frame, text="Select Voice:")
        self.voice_label.grid(row=0, column=0, sticky=tk.W)

        self.voice_var = tk.StringVar()
        self.voice_var.set("default")

        voices = self.engine.getProperty('voices')
        num_voices = len(voices)
        num_columns = 10
        num_rows = (num_voices + num_columns - 1) // num_columns

        for i, voice in enumerate(voices):
            voice_radio = tk.Radiobutton(self.options_frame, text=voice.name, variable=self.voice_var, value=voice.id)
            voice_radio.grid(row=(i // num_columns) + 1, column=i % num_columns, sticky=tk.W)

        self.button_frame = tk.Frame(self.frame)
        self.button_frame.pack(pady=10)

        self.select_button = tk.Button(self.button_frame, text="Select PDF", command=self.select_pdf)
        self.select_button.pack(side=tk.LEFT, padx=5)

        self.read_button = tk.Button(self.button_frame, text="Read", command=self.start_reading)
        self.read_button.pack(side=tk.LEFT, padx=5)

        self.save_button = tk.Button(self.button_frame, text="Save Audio", command=self.save_audio)
        self.save_button.pack(side=tk.LEFT, padx=5)

        self.exit_button = tk.Button(self.button_frame, text="Exit", command=self.window.quit)
        self.exit_button.pack(side=tk.LEFT, padx=5)

        self.reading_started = False

        self.text_display.bind("<ButtonRelease-1>", self.on_text_selection)

    def extract_text_from_pdf(self):
        with open(self.file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfFileReader(file)
            text = ""
            for page in range(pdf_reader.numPages):
                page_content = pdf_reader.getPage(page).extract_text()
                text += page_content
            return text

    def start_reading(self):
        if not self.reading_started:
            if not self.selected_text:
                messagebox.showinfo("No Selection", "Please select the text you want to read.")
                return
            self.reading_started = True
            self.engine.connect('started-utterance', self.on_start_utterance)
            self.engine.connect('finished-utterance', self.on_end_utterance)

            Thread(target=self.read_text).start()

    def read_text(self):
        self.engine.setProperty('voice', self.voice_var.get())
        self.engine.say(self.selected_text)
        self.engine.runAndWait()
        self.reading_started = False

    def on_start_utterance(self, name):
        self.text_display.tag_config("highlight", background="yellow")
        self.text_display.tag_add("highlight", "sel.first", "sel.last")
        self.text_display.see("insert")

    def on_end_utterance(self, name, completed):
        self.text_display.tag_remove("highlight", "sel.first", "sel.last")

    def on_text_selection(self, event):
        if self.text_display.tag_ranges("sel"):
            self.selected_text = self.text_display.selection_get()

    def select_pdf(self):
        file_path = filedialog.askopenfilename(title="Select PDF file", filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.file_path = file_path
            self.text = self.extract_text_from_pdf()
            self.text_display.delete(1.0, tk.END)
            self.text_display.insert(tk.END, self.text)

    def save_audio(self):
        if not self.selected_text:
            messagebox.showinfo("No Selection", "Please select the text you want to save as audio.")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".mp3", filetypes=[("MP3 Files", "*.mp3")])
        if save_path:
            self.engine.setProperty('voice', self.voice_var.get())
            self.engine.save_to_file(self.selected_text, save_path)
            self.engine.runAndWait()

    def run(self):
        self.window.mainloop()


if __name__ == "__main__":
    reader = PDFReader()
    reader.run()
