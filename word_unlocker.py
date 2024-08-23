import zipfile
import os
import shutil
import xml.etree.ElementTree as ET


def unlock_docx_files(input_dir, output_dir):
    # Ensure the output directory exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Process each .docx file in the input directory
    for filename in os.listdir(input_dir):
        if filename.endswith(".docx"):
            docx_path = os.path.join(input_dir, filename)
            output_path = os.path.join(output_dir, filename)

            # Step 1: Create a temporary directory to extract the .docx contents
            temp_dir = os.path.join(output_dir, "temp_docx")
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            os.makedirs(temp_dir)

            # Step 2: Extract the .docx file
            with zipfile.ZipFile(docx_path, 'r') as docx_zip:
                docx_zip.extractall(temp_dir)

            # Step 3: Locate and modify settings.xml to remove protection
            settings_path = os.path.join(temp_dir, 'word', 'settings.xml')
            if os.path.exists(settings_path):
                tree = ET.parse(settings_path)
                root = tree.getroot()

                # Namespace dictionary (needed due to the XML namespaces used in .docx files)
                ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

                # Remove documentProtection element if it exists
                protection_element = root.find('w:documentProtection', ns)
                if protection_element is not None:
                    root.remove(protection_element)
                    print(f"Removed document protection for {filename}.")

                # Write the modified XML back to settings.xml
                tree.write(settings_path)

            # Step 4: Repackage the .docx file
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx_zip:
                for foldername, subfolders, filenames in os.walk(temp_dir):
                    for filename in filenames:
                        file_path = os.path.join(foldername, filename)
                        arcname = os.path.relpath(file_path, temp_dir)
                        docx_zip.write(file_path, arcname)

            # Step 5: Clean up the temporary directory
            shutil.rmtree(temp_dir)

            print(f"Unlocked document saved as {output_path}")

# Example usage with specified paths:
unlock_docx_files(r"C:\Users\franc\Desktop\docx", r"C:\Users\franc\Desktop\docx output")
