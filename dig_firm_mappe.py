import pandas as pd
import os
import shutil  # Import shutil for file operations

# Read the first sheet of an Excel file
df = pd.read_excel('./hoko-ausstellerliste-43-2024-10-05-17-12-55.xlsx')

# Create the 'Folder' column by joining the specified columns
df['Folder'] = df['Standinformationen'].str[1:3] + '_' + df['Standnumber'].astype(str) + '_' + df['Firma']

# Define forbidden characters
forbidden_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']

# Path to the PDF files that need to be copied
pdf_files = {
    'VDE': r"C:\Users\justi\PycharmProjects\firmenMappe\PdfDateien\VDE-Erklärung x 100 A4_2024.pdf",
    'Datenschutz': r"C:\Users\justi\PycharmProjects\firmenMappe\PdfDateien\Datenschutzerklärung_2024.pdf"
}

# Base path for folder creation
base_path = r"C:\Users\justi\PycharmProjects\firmenMappe\HOKO_2024"

# Function to sanitize folder names
def sanitize_folder_name(folder_name, forbidden_chars):
    for char in forbidden_chars:
        folder_name = folder_name.replace(char, '_')
    return folder_name

# Function to create directory and copy PDF files
def create_folder_and_copy_pdfs(folder, parent_folder, pdf_files):
    # Define the full path where you want to create the directories
    folder_path = os.path.join(base_path, parent_folder, folder)

    # Create the parent and subdirectory if they don't already exist
    os.makedirs(folder_path, exist_ok=True)

    # Copy each PDF file into the newly created folder
    for pdf_name, pdf_path in pdf_files.items():
        pdf_destination = os.path.join(folder_path, os.path.basename(pdf_path))
        try:
            shutil.copy(pdf_path, pdf_destination)
            print(f"Copied {pdf_name} PDF to: {pdf_destination}")
        except Exception as e:
            print(f"Error copying {pdf_name} PDF to {folder_path}: {e}")

# Create directories and copy PDFs based on the 'Folder' column
for folder in df["Folder"]:
    # Check for NaN and convert to string if necessary
    if pd.isna(folder):
        folder = 'Default_Folder'  # Provide a default name for NaN values
    else:
        folder = sanitize_folder_name(str(folder), forbidden_chars)

    # Determine the parent directory based on the first two characters
    if folder.startswith("Di"):
        parent_folder = "Dienstag"
    elif folder.startswith("Mi"):
        parent_folder = "Mittwoch"
    elif folder.startswith("Do"):
        parent_folder = "Donnerstag"
    else:
        parent_folder = "Other"

    # Only create folder and copy PDFs if in 'Dienstag', 'Mittwoch', or 'Donnerstag'
    if parent_folder in ["Dienstag", "Mittwoch", "Donnerstag"]:
        create_folder_and_copy_pdfs(folder, parent_folder, pdf_files)
