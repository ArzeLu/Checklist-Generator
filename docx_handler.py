from docx import Document # type: ignore (warning suppressant for vscode that doesn't have the library globally)
from docx.shared import Pt # type: ignore (warning suppressant for vscode that doesn't have the library globally)
from tkinter import *
from tkinter import messagebox
import sys
import os

# local imports
from log_handler import TextFileHandler

class DocxHandler():
    def __init__(self, root):
        self.root = root
        self.txt_handler = TextFileHandler(self.root)
        self.serial_placeholder = "{serial number}"
        self.start_time_placeholder = "{start time}"
        self.end_time_placeholder = "{end time}"

    ## From source file to the checklist:
    ## Fill in serial number
    ## Fill in burn in times
    ## Checks if any placeholder is missing from the checklist and throw appropriate exceptions
    ##
    ## Arguments: 1. The checklist template, 2. One target file name
    def fill_checklist_info(self, new_checklist, file_name, source_file_dir):
        paragraphs = new_checklist.paragraphs
        # fill in info
        serial, start_time, end_time = self.txt_handler.find_info(file_name, source_file_dir)

        if serial is None:
            error_message = f"The burn in file {file_name} has incomplete data\n\nTerminating program\n\nThe generated ones are still there, go check them"
            messagebox.showinfo(title = "HAHA REBURN!!!!!", message = error_message, icon = "error", detail = "Gotta ask Brandon for forgiveness now...")
            return None

        serial_placeholder_found = False
        start_time_placeholder_found = False
        end_time_placeholder_found = False

        for paragraph in paragraphs:
            runs = paragraph.runs
            for run in runs:
                if self.serial_placeholder in run.text:
                    run.text = run.text.replace(self.serial_placeholder, serial)
                    serial_placeholder_found = True

                if self.start_time_placeholder in run.text:
                    run.text = run.text.replace(self.start_time_placeholder, start_time)
                    start_time_placeholder_found = True
                    continue

                if self.end_time_placeholder in run.text:
                    run.text = run.text.replace(self.end_time_placeholder, end_time)
                    end_time_placeholder_found = True
                    continue

        # Check for any missing placeholders
        error = False

        if not serial_placeholder_found:
            messagebox.showinfo(title = "AYO gotta fix your serial numbers dawg", message = "Checklist template doesn't have {serial number} placeholder!", icon = "error", detail = "Michael is going to steal your doge coins!")
            error = True

        if not start_time_placeholder_found:
            messagebox.showinfo(title = "the procrastination when im writing this code is real", message = "Checklist template doesn't have {start time} placeholder!", icon = "error", detail = "Hurry up before Stanley notices!")
            error = True

        if not end_time_placeholder_found:
            messagebox.showinfo(title = "the end is never the end is never the end is never the end", message = "Checklist template doesn't have {end time} placeholder!", icon = "error", detail = "better enter an end time! KC doesn't like overtime...watch your back...")
            error = True

        if error:
            return None
        
        return new_checklist

    ## Generate checklists with appropriate fields
    ## Arguments: 1. directory of the checklist, 2. directory of the source files, 3. directory of generated files
    ## Only checking the validity of directories here because it's the only time when they're used.
    def generate_checklists(self, checklist_dir, source_files_dir, destination_dir):
        source_file_names = self.txt_handler.get_file_names(source_files_dir)

        error = False

        # check if the directories are valid.
        # use exists for both files and directories
        if os.path.exists(checklist_dir) is not True:
            messagebox.showinfo(title = "wah wah wrong file", message = "Invalid Checklist File!", icon = "error", detail = "im gonna call the IPQC on you")
            error = True

        if os.path.isdir(source_files_dir) is not True:
            messagebox.showinfo(title = "checking in to see if the vga still works", message = "Invalid Source File Directory", icon = "error", detail = "lets hope you didn't lose the files")
            error = True

        if os.path.isdir(destination_dir) is not True:
            messagebox.showinfo(title = "batch one...batch two...b- nevermind", message = "Invalid Destination Folder!", icon = "error", detail = "Ask Edwin for a new dos bruh. jks he's not gonna give it to you")
            error = True
        
        if error:
            return
        
        count = 0
        for file_name in source_file_names:
            new_checklist = self.fill_checklist_info(Document(checklist_dir), file_name, source_files_dir)
            if new_checklist is None:
                return
            new_checklist.save(destination_dir + "/" + file_name + ".docx")
            count += 1

        messagebox.showinfo(title = "yayayayayayya", message = "Successfully generated all checklists", icon = "info", detail = "time to pack it all up...")