# TODO: replace sys exit warnings with pop up windows.

from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import time
import random

# local imports
from docx_handler import DocxHandler

APP_VERSION = "2.0.0"        

class UserInterface():
    def __init__(self):
        self.root = Tk()
        self.doc_handler = DocxHandler(self.root)
        self.source_files_dir = StringVar(self.root)
        self.checklist_template_dir = StringVar(self.root)
        self.destination_dir = StringVar(self.root)
        self.directory_entries = [self.source_files_dir, self.checklist_template_dir, self.destination_dir]
        self.emotes = ['📁', '👽', '🐱', '🦷', '🎁', '👾', '🦜', '💀', '💩', '🍥', '🥶', '💅', '🍕', '🚗', '🎃']
        self.emotes_count = 15
    
    def select_burnin_dir(self):
        self.source_files_dir.set(filedialog.askdirectory(initialdir = "/", title = "🔥🔥🔥🔥🔥"))

    def select_checklist_dir(self):
        self.checklist_template_dir.set(filedialog.askopenfilename(initialdir = "/", title = "📄📄📄📄📄", filetypes = [("💋💋💋💋💋💋💋💋", "*.*")]))

    def select_destination_dir(self):
        self.destination_dir.set(filedialog.askdirectory(initialdir = "/", title = "🎯🎯🎯🎯🎯"))

    def choose_emote(self):
        n = random.randrange(0, self.emotes_count)  # first argument is included, not the second.
        return self.emotes[n]

    def generate_button_action(self):
        self.doc_handler.generate_checklists(self.checklist_template_dir.get(), self.source_files_dir.get(), self.destination_dir.get())

    def reroll(self):
        for entry in self.directory_entries:
            entry.set("")
	
        self.root.update()  # This bypasses the buffer for the gui
        time.sleep(0.3)

        for entry in self.directory_entries:
            entry.set("       --> " + self.choose_emote() + " <--")
            self.root.update()
            time.sleep(0.2)

    def run(self):
        self.root.title("Checklist Generator")
        self.root.geometry("600x200")
        self.root.columnconfigure(0, weight = 1)
        self.root.rowconfigure(0, weight = 1)
        self.root.rowconfigure(1, weight = 1)
        self.root.resizable(False, False)

        s = ttk.Style()

        ## Elements of the first three rows ##
        s.configure("Top.TFrame")
        top_frame = ttk.Frame(self.root, padding = (10, 20, 5, 5), style = "Top.TFrame")
        top_frame.grid(column = 0, row = 0, sticky = N+S+W+E)  # set the starting position of the top_frame

        burnin_label = ttk.Label(top_frame, text = "Burn-in files: ", font = (15))
        checklist_label = ttk.Label(top_frame, text = "Checklist template: ", font = (15))
        destination_label = ttk.Label(top_frame, text = "Destination folder:", font = (15))

        burnin_dir = ttk.Entry(top_frame, textvariable = self.source_files_dir, width = 25, background = "white", font = (15))
        checklist_dir = ttk.Entry(top_frame, textvariable = self.checklist_template_dir, width = 25, background = "white", font = (15))
        destination_dir = ttk.Entry(top_frame, textvariable = self.destination_dir, width = 25, background = "white", font = (15))

        burnin_button = ttk.Button(top_frame, text = "Browse", command = self.select_burnin_dir)
        checklist_button = ttk.Button(top_frame, text = "Browse", command = self.select_checklist_dir)
        destination_button = ttk.Button(top_frame, text = "Browse", command = self.select_destination_dir)

        burnin_label.grid(column = 0, row = 0, columnspan = 3, padx = (5, 5), pady = (0, 5), sticky = N+S+W+E)
        checklist_label.grid(column = 0, row = 1, columnspan = 3, padx = (5, 5), pady = (0, 5), sticky = N+S+W+E)
        destination_label.grid(column = 0, row = 2, columnspan = 3, padx = (5, 5), pady = (0, 5), sticky = N+S+W+E)

        random.seed(time.time_ns())  # time_ns() instead of time() because nanoseconds better
        # self.reroll()
        burnin_dir.grid(column = 3, row = 0, padx = (0, 1), sticky = N+S+W+E)
        checklist_dir.grid(column = 3, row = 1, padx = (1, 1), sticky = N+S+W+E)
        destination_dir.grid(column = 3, row = 2, padx = (1, 1), sticky = N+S+W+E)

        burnin_button.grid(column = 4, row = 0, padx = (5, 5), pady = (0, 5), sticky = N+S+W+E)
        checklist_button.grid(column = 4, row = 1, padx = (5, 5), pady = (0, 5), sticky = N+S+W+E)
        destination_button.grid(column = 4, row = 2, padx = (5, 5), pady = (0, 5), sticky = N+S+W+E)

        ## Remaining buttons on the bottom half of the screen ##
        s.configure("Bottom.TFrame")
        bottom_frame = ttk.Frame(self.root, padding = (5, 5, 5, 5), style = "Bottom.TFrame")
        bottom_frame.grid(column = 0, row = 1, sticky = N+S+W+E)

        # Function buttons
        s.configure("Button.TButton", font = (15))

        insert_button = ttk.Button(bottom_frame, text = "Insert Serial Numbers", style = "Button.TButton")
        insert_button.grid(column = 0, row = 0, padx = (10, 10), pady = (10, 10), sticky = N+S+W+E)

        generate_button = ttk.Button(bottom_frame, text = "Generate Checklists", style = "Button.TButton", command = self.generate_button_action)
        generate_button.grid(column = 1, row = 0, padx = (10, 10), pady = (10, 10), sticky = N+S+W+E)    

        gamble_button = ttk.Button(bottom_frame, text = "Reroll", style = "Button.TButton", command = self.reroll)
        gamble_button.grid(column = 2, row = 0, padx = (10, 10), pady = (10, 10), sticky = N+S+W+E)

        self.root.mainloop()

class Main():
    def __init__(self):
        self.gui = UserInterface()
    def run(self):
        self.gui.run()
        
if __name__ == "__main__":
    driver = Main()
    driver.run()