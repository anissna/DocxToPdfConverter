import os
import comtypes.client

# Set the path of the folder containing the .docx files
folder_path = r"C:\Users\user\OneDrive\Bureau\2023"

# Create a list of all .docx files in the folder
docx_files = [f for f in os.listdir(folder_path) if f.endswith(".docx")]

# Create a subfolder to store the converted .pdf files
pdf_folder = os.path.join(folder_path, "PDF_Files")
os.makedirs(pdf_folder, exist_ok=True)

# Loop through each .docx file and convert it to .pdf
for docx_file in docx_files:
    # Set the input and output paths
    input_path = os.path.join(folder_path, docx_file)
    output_path = os.path.join(pdf_folder, docx_file.replace(".docx", ".pdf"))
    
    # Try to open the .docx file and save it as .pdf
    try:
        # Create a new Word application instance
        word = comtypes.client.CreateObject("Word.Application")

        # Disable pop-up alerts
        word.DisplayAlerts = False

        # Open the .docx file and save it as .pdf
        doc = word.Documents.Open(input_path)
        doc.SaveAs(output_path, FileFormat=17)

        # Close the document and quit Word
        doc.Close()
        word.Quit()

        # Let the user know that the file has been successfully converted
        print("[☠ビトレス ☠]: The file " + docx_file + " has been converted to .pdf and saved in the PDF_Files folder.")

    except Exception as e:
        # If there is an error, print a message indicating that the file could not be converted
        print("[☠ビトレス ☠]: The file " + docx_file + " could not be converted to .pdf. Error message: " + str(e))

# Let the user know the conversion is complete
print("[☠ビトレス ☠]: Conversion complete.")
