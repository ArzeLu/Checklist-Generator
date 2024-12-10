import os
import sys
import re
from tkinter import messagebox

# local imports
from helper import Helper

class TextFileHandler():
    def __init__(self, root):
        self.root = root
        self.helper = Helper()
        self.source_files_dir = ""

    def get_file_names(self, source_files_dir):
        self.source_files_dir = source_files_dir
        source_file_names = []
        clean_file_names = []

        # Get burn in file names 
        for root, dirs, files in os.walk(source_files_dir):
            source_file_names.extend(files)

        # Remove the file extensions from the file name strings
        for file in source_file_names:
            clean_file_names.append(os.path.splitext(file)[0])

        return clean_file_names

    def find_info(self, file_name, source_file_dir):
        serial = file_name
        start_time = ""
        end_time = ""

        # find burn in times from the burn in .txt
        dates = []
        with open(source_file_dir + "/" + file_name + ".log", encoding = "utf_16") as file:
            lines = file.readlines()

            for line in lines:
                line = line.rstrip()
                date_match = re.search(r"\d{4}[-]\d{2}[-]\d{2}\s\d{2}[:]\d{2}[:]\d{2}", line)
                if date_match is not None:
                    dates.append(date_match.group())

        start_time = None
        end_time = None

        try:
            # there will be three matches, only get the first and last one.
            # look at a burn-in report for reference
            start_time = dates[0]
            end_time = dates[2]

            start_time = self.helper.convert_timezone(start_time)
            end_time = self.helper.convert_timezone(end_time)

        except:
            return [None, None, None]

        return [serial, start_time, end_time]