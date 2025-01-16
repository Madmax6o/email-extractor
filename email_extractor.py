import os
import re
import pandas as pd
import docx
import zipfile
import time
from typing import List
from tkinter import Tk, Label, Button, Entry, filedialog, IntVar, Radiobutton, StringVar

def extract_emails_from_text(text: str, domain_filters: List[str]) -> List[str]:
    email_pattern = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
    all_emails = email_pattern.findall(text)
    if domain_filters:
        return [email for email in all_emails if any(email.endswith(domain) for domain in domain_filters)]
    return all_emails

def extract_emails_from_file(file_path: str, domain_filters: List[str]) -> List[str]:
    emails = []
    _, file_extension = os.path.splitext(file_path)

    try:
        if file_extension in ['.txt', '.csv', '.sql']:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                content = file.read()
                emails.extend(extract_emails_from_text(content, domain_filters))
        elif file_extension in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path, engine='openpyxl')
            for column in df.columns:
                emails.extend(extract_emails_from_text(' '.join(df[column].astype(str)), domain_filters))
        elif file_extension == '.docx':
            doc = docx.Document(file_path)
            for paragraph in doc.paragraphs:
                emails.extend(extract_emails_from_text(paragraph.text, domain_filters))
        elif file_extension == '.zip':
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                # Extract and process only files named "valid.csv"
                for file_name in zip_ref.namelist():
                    if file_name.endswith('valid.csv'):
                        with zip_ref.open(file_name) as csv_file:
                            content = csv_file.read().decode('utf-8', errors='ignore')
                            emails.extend(extract_emails_from_text(content, domain_filters))
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")

    return emails

def is_hidden(filepath):
    return bool(os.stat(filepath).st_file_attributes & (0x2 | 0x4 | 0x8))

def search_emails_in_directory(directory_path: str, domain_filters: List[str], progress_var: IntVar, progress_label: StringVar, root: Tk) -> List[str]:
    all_emails = []
    files_list = []
    for root_dir, dirs, files in os.walk(directory_path):
        # Skip hidden directories
        dirs[:] = [d for d in dirs if not is_hidden(os.path.join(root_dir, d))]
        for file in files:
            if file.endswith(('.txt', '.csv', '.sql', '.xlsx', '.xls', '.docx', '.zip')):
                files_list.append(os.path.join(root_dir, file))
    
    total_files = len(files_list)
    for idx, file_path in enumerate(files_list):
        # Skip hidden files
        if not is_hidden(file_path):
            all_emails.extend(extract_emails_from_file(file_path, domain_filters))
        # Update progress
        progress = int((idx + 1) / total_files * 100)
        progress_var.set(progress)
        progress_label.set(f"Progress: {progress}%")
        root.update_idletasks()

    return list(set(all_emails))

def save_emails_to_file(emails: List[str], output_file: str):
    with open(output_file, 'w') as file:
        for email in emails:
            file.write(email + '\n')

def run_email_extraction(option, path, output, domain_filters, progress_var, progress_label, time_label, root):
    start_time = time.time()
    emails = []
    if option == 1:
        emails = search_emails_in_directory(path, domain_filters, progress_var, progress_label, root)
    else:
        print("Invalid option selected.")
        return

    if emails:
        save_emails_to_file(emails, output)
        print(f"Found {len(emails)} unique email addresses. Saved to {output}.")
    else:
        print("No email addresses found.")
    
    end_time = time.time()
    elapsed_time = end_time - start_time
    time_label.set(f"Elapsed Time: {elapsed_time:.2f} seconds")

def select_path():
    path = filedialog.askdirectory()
    path_entry.delete(0, 'end')
    path_entry.insert(0, path)

def select_output():
    output = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
    output_entry.delete(0, 'end')
    output_entry.insert(0, output)

def execute():
    option = option_var.get()
    path = path_entry.get()
    output = output_entry.get()
    domain_filters = [domain.strip() for domain in domain_entry.get().split(',')]
    run_email_extraction(option, path, output, domain_filters, progress_var, progress_label, time_label, root)

# GUI Setup
root = Tk()
root.title("Email Extractor")

option_var = IntVar(value=1)
progress_var = IntVar()
progress_label = StringVar()
time_label = StringVar()

Label(root, text="Path:").grid(row=0, column=0, sticky='w')
path_entry = Entry(root, width=50)
path_entry.grid(row=0, column=1)
Button(root, text="Browse", command=select_path).grid(row=0, column=2)

Label(root, text="Output file:").grid(row=1, column=0, sticky='w')
output_entry = Entry(root, width=50)
output_entry.grid(row=1, column=1)
Button(root, text="Browse", command=select_output).grid(row=1, column=2)

Label(root, text="Email domain filters (comma-separated):").grid(row=2, column=0, sticky='w')
domain_entry = Entry(root, width=50)
domain_entry.grid(row=2, column=1)

Label(root, textvariable=progress_label).grid(row=3, column=0, columnspan=3)
Label(root, textvariable=time_label).grid(row=4, column=0, columnspan=3)

Button(root, text="Run", command=execute).grid(row=5, column=0, columnspan=3)

root.mainloop()
