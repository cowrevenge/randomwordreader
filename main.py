import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import random
import pyttsx3

# Initialize text-to-speech engine
engine = pyttsx3.init()


class RandomWordSelectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Random Word Selector")
        self.root.geometry("400x400")

        # Load button
        self.load_button = tk.Button(root, text="Load Excel Sheet", command=self.load_sheet)
        self.load_button.pack(pady=10)

        # Select button
        self.select_button = tk.Button(root, text="Select Random Word", command=self.select_word)
        self.select_button.pack(pady=10)
        self.select_button["state"] = "disabled"  # Initially disabled

        # Read Word button
        self.read_button = tk.Button(root, text="Read Word", command=self.read_word)
        self.read_button.pack(pady=10)
        self.read_button["state"] = "disabled"  # Initially disabled

        # Label to display selected word
        self.word_label = tk.Label(root, text="", font=("Helvetica", 16))
        self.word_label.pack(pady=20)

        # Log of the last 10 words
        self.log_label = tk.Label(root, text="Last 10 Words:", font=("Helvetica", 12))
        self.log_label.pack()

        self.log_box = tk.Listbox(root, height=10, width=40, font=("Helvetica", 10))
        self.log_box.pack(pady=10)
        self.log_box.bind("<Double-Button-1>", self.on_double_click)

        self.data = None  # To store the words from the Excel file
        self.current_word = None  # To store the currently selected word
        self.word_log = []  # To store the last 10 words

    def load_sheet(self):
        # Ask for the Excel file path
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                # Read the Excel file
                self.data = pd.read_excel(file_path, header=None)
                if not self.data.empty:
                    messagebox.showinfo("Success", "Excel sheet loaded successfully!")
                    self.select_button["state"] = "normal"  # Enable the select button
                else:
                    messagebox.showerror("Error", "The Excel sheet is empty.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel sheet: {str(e)}")

    def select_word(self):
        if self.data is not None:
            # Select a random word
            random_row = self.data.sample().values.flatten()
            random_word = random.choice(random_row)

            # Update the label with the random word
            self.word_label.config(text=random_word)

            # Store the current word
            self.current_word = random_word

            # Enable the read button
            self.read_button["state"] = "normal"

            # Update the log with FIFO (keep the last 10 words)
            self.update_log(random_word)

    def read_word(self):
        if self.current_word:
            # Disable the read button to avoid queuing multiple clicks
            self.read_button["state"] = "disabled"

            # Speak the current word
            engine.say(self.current_word)
            engine.runAndWait()

            # Wait for .5 second before enabling the read button again
            self.root.after(500, self.enable_read_button)

    def enable_read_button(self):
        self.read_button["state"] = "normal"

    def update_log(self, word):
        # Append new word to the log
        if len(self.word_log) >= 10:
            # Remove the oldest word if we have more than 10
            self.word_log.pop(0)
        self.word_log.append(word)

        # Update the log box in the GUI
        self.log_box.delete(0, tk.END)  # Clear current log display
        for w in self.word_log:
            self.log_box.insert(tk.END, w)  # Insert each word into the log box

    def on_double_click(self, event):
        # Get the clicked word from the log box
        selected_index = self.log_box.curselection()
        if selected_index:
            selected_word = self.log_box.get(selected_index)
            self.word_label.config(text=selected_word)  # Display it in the label
            self.current_word = selected_word  # Set it as the current word
            self.read_button["state"] = "normal"  # Enable the read button


# Create the main window
root = tk.Tk()
app = RandomWordSelectorApp(root)
root.mainloop()
