import os
import win32com.client


def convert_docx_to_pdf(input_folder, output_folder):
    # Create a new instance of Word application
    word = win32com.client.Dispatch("Word.Application")

    # Iterate through all files in the input folder
    for filename in os.listdir(input_folder):
        if filename.endswith(".docx"):
            try:
                # Construct full paths for input and output files
                input_path = os.path.join(input_folder, filename)
                # Remove the .docx extension and add .pdf
                output_path = os.path.join(
                    output_folder, filename[:-5] + ".pdf")

                # Open the Word document
                doc = word.Documents.Open(input_path)

                # Save the document as PDF
                # FileFormat 17 represents PDF format
                doc.SaveAs(output_path, FileFormat=17)

                # Close the document
                doc.Close()

                # DELETE
                # os.remove(input_path)
            except Exception:
                print("An error happened")
                continue

    # Quit Word application
    word.Quit()


# Get current working directory
current_dir = os.getcwd()

# Use current directory as input and output folders
convert_docx_to_pdf(current_dir, current_dir)
