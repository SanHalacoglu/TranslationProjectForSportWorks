import os
import shutil
from tkinter import filedialog, messagebox

import openpyxl
from openai import OpenAI


def is_valid_openai_api_key(api_key: str):
    api_key = api_key.strip()
    if api_key.startswith("sk-proj-"):
        return len(api_key) <= 300
    return False

def save_file():
    target_file = filedialog.asksaveasfilename(
        title="Save Excel File As",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )

def save_excel_template(source_file: str, default_filename: str = "TemplateForTranslation.xlsx") -> None:
    if not os.path.exists(source_file):
        messagebox.showerror("File Not Found", f"The file '{source_file}' does not exist.")
        return

    target_file = filedialog.asksaveasfilename(
        title="Save Excel File As",
        defaultextension=".xlsx",
        initialfile=default_filename,
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )

    if not target_file:
        return

    try:
        shutil.copy(source_file, target_file)
        messagebox.showinfo("Download Complete", f"The file has been saved as '{target_file}'.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while saving the file: {e}")

def save_excel_workbook(workbook: openpyxl.Workbook, default_filename: str = "TranslatedFile.xlsx") -> str:
    target_file = filedialog.asksaveasfilename(
        title="Save Translated Excel File As",
        defaultextension=".xlsx",
        initialfile=default_filename,
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )

    if not target_file:
        return ""

    try:
        workbook.save(target_file)
        messagebox.showinfo("Download Complete", f"The file has been saved as '{target_file}'.")
        return target_file
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while saving the file: {e}")
        return ""


