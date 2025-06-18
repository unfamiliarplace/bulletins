# TODO Add name detection for the [Commenter: Remark] style ones and fit them in somewhere?

from __future__ import annotations

import os
from pathlib import Path
from tkinter import Tk, messagebox
from tkinter.filedialog import askdirectory
import traceback

from reports import Reports

PATH_OUTPUT_SUFFIX = 'collated'
PATH_SRC = Path('src')
PATH_INPUT_DEFAULT = PATH_SRC / 'input'
PATH_OUTPUT_DEFAULT = PATH_INPUT_DEFAULT / PATH_OUTPUT_SUFFIX

def prompt_directory(initialdir: Path|str=os.getcwd()) -> str:
    pathstr = askdirectory(title='Input directory', initialdir=initialdir)
    return Path(pathstr)

def generate_reports() -> None:
    path_input = prompt_directory(PATH_INPUT_DEFAULT)
    path_output = path_input / PATH_OUTPUT_SUFFIX

    filenames = Reports.list_files(path_input)
    dses = Reports.parse_files(filenames)
    aliases = Reports.map_aliases(path_input)

    students = Reports.collate_students(dses, aliases)
    documents = Reports.make_documents(students)
    Reports.save_documents(documents, path_output)

if __name__ == '__main__':
    root = Tk()
    root.title('TDChristian Personality Questionnaire')
    try:
        generate_reports()
    except Exception as e:
        messagebox.showerror(title='Error', message=f'Could not process.\n{repr(e)}\n\n{traceback.format_exc()}')
    finally:
        root.destroy()
