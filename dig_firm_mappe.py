import pandas as pd
import os

# Read the first sheet of an Excel file
df = pd.read_excel('../hoko-ausstellerliste-43-2024-10-05-17-12-55.xlsx')

# Create the 'Folder' column by joining the specified columns
df['Folder'] = df['Standinformationen'].str[1:3] + '_' + df['Standnumber'].astype(str) + '_' + df['Firma']

# Define forbidden characters
forbidden_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']

# Create directories based on the 'Folder' column
for folder in df["Folder"]:
    # Check for NaN and convert to string if necessary
    if pd.isna(folder):
        folder = 'Default_Folder'  # Provide a default name for NaN values
    else:
        folder = str(folder)  # Ensure folder is a string

    # Replace each forbidden character with '_'
    for char in forbidden_chars:
        folder = folder.replace(char, '_')

    # Determine the parent directory based on the first two characters
    if folder.startswith("Di"):
        parent_folder = "Dienstag"
    elif folder.startswith("Mi"):
        parent_folder = "Mittwoch"
    elif folder.startswith("Do"):
        parent_folder = "Donnerstag"
    else:
        parent_folder = "Other"  # Use a default parent for other cases

    # Define the full path where you want to create the directories
    base_path = r"C:\Users\lutsc\PycharmProjects\digitale_firmenmappe\HOKO_2024"
    parent_path = os.path.join(base_path, parent_folder)
    folder_path = os.path.join(parent_path, folder)

    # Create the parent directory if it doesn't already exist
    os.makedirs(parent_path, exist_ok=True)

    # Create the directory (if it doesn't already exist)
    try:
        os.mkdir(folder_path)
        print(f"Created directory: {folder_path}")
    except FileExistsError:
        print(f"Directory already exists: {folder_path}")
    except Exception as e:
        print(f"Error creating directory {folder_path}: {e}")
