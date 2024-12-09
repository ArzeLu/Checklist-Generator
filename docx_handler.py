# TODO: replace sys exit warnings with pop up windows.

from docx import Document # type: ignore (warning suppressant for vscode that doesn't have the library globally)
from docx.shared import Pt # type: ignore (warning suppressant for vscode that doesn't have the library globally)
from tkinter import *
import sys
import os

# local imports
from log_handler import TextFileHandler

class DocxHandler():
    def __init__(self):
        self.txt_handler = TextFileHandler()
        self.serial_placeholder = "{serial number}"
        self.start_time_placeholder = "{start time}"
        self.end_time_placeholder = "{end time}"

    ## From source file to the checklist:
    ## Fill in serial number
    ## Fill in burn in times
    ## Checks if any placeholder is missing from the checklist and throw appropriate exceptions
    ##
    ## Arguments: 1. The checklist template, 2. One target file name
    def fill_checklist_info(self, checklist, file_name, source_file_dir):
        new_checklist = checklist
        paragraphs = new_checklist.paragraphs
        tables = new_checklist.tables
        # fill in info
        serial, start_time, end_time = self.txt_handler.find_info(file_name, source_file_dir)
        serial_placeholder_found = False
        start_time_placeholder_found = False
        end_time_placeholder_found = False

        for paragraph in paragraphs:
            runs = paragraph.runs
            for run in runs:
                print(run.text)
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
            print("""
            =============== Error! =================\n
            Checklist template doesn't have {serial number} placeholder!\n
            ========================================\n
            """)
            error = True

        if not start_time_placeholder_found:
            print("""
            =============== Error! =================\n
            Checklist template doesn't have {start time} placeholder!\n
            ========================================\n
            """)
            error = True

        if not end_time_placeholder_found:
            print("""
            =============== Error! =================\n
            Checklist template doesn't have {end time} placeholder!\n
            ========================================\n
            """)
            error = True

        if error:
            sys.exit()

        return new_checklist

    ## Generate checklists with appropriate fields
    ## Arguments: 1. directory of the checklist, 2. directory of the source files, 3. directory of generated files
    ## Only checking the validity of directories here because it's the only time when they're used.
    def generate_checklists(self, checklist_dir, source_files_dir, destination_dir):
        source_file_names = self.txt_handler.get_file_names(source_files_dir)
        checklist = Document(checklist_dir)

        # check if the directories are valid.
        # use exists for both files and directories
        error = False
        if os.path.exists(checklist_dir) is not True:
            print("""
            =============== Error! =================\n
            Invalid checklist directory!\n
            ========================================\n
            """)
            error = True

        if os.path.isdir(source_files_dir) is not True:
            print("""
            =============== Error! =================
            Invalid source files directory!\n
            ========================================\n
            """)
            error = True

        if os.path.isdir(destination_dir) is not True:
            print("""
            =============== Error! =================\n
            Invalid destination directory!\n
            ========================================\n
            """)
            error = True

        if error:
            sys.exit()
        
        for file_name in source_file_names:
            new_checklist = self.fill_checklist_info(checklist, file_name, source_files_dir)
            new_checklist.save(destination_dir + "/" + file_name + ".docx")