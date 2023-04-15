import os
import random
import fitz  # pip install PyMuPDF

folder_path = ("Your/Path/To/All/Files")        # Add your path
required_pages = 385


def main(folder_path: str, required_pages: int):
    """
    This function selects random pages from PDF files in a specified folder and saves them as PNG images.

    Parameters:
    folder_path (str): The path to the folder containing the PDF files.
    required_pages (int): The number of pages to select.
    """
    # Get a list of all the files in the folder
    files = os.listdir(folder_path)
    files_count = len(files)

    # Create a new folder to store the selected pages as images
    if not os.path.exists(os.path.join(folder_path, "images")):
        os.mkdir(os.path.join(folder_path, "images"))
    
    # Select random pages and save them as PNG images
    count = 0
    while count < required_pages:

        try:
            # Choose a random file from the list
            rand_file = random.randint(0, files_count - 1)
            cur_file = os.path.join(folder_path, files[rand_file])

            # Open the PDF file and choose a random page
            doc = fitz.open(cur_file)
            rand_page = random.randint(0, len(doc) - 1)
            page = doc.load_page(rand_page)
            pix = page.get_pixmap()

            # Save the page as a PNG image
            output_name = f"{files[rand_file][:-4]}-{rand_page}.png"
            output_file_path = os.path.join(folder_path, "images", output_name)
            pix.save(output_file_path)
            doc.close()
            count += 1
        except:
            pass
    print("Pages were saved successfully!")


if __name__ == "__main__":
    main(folder_path, required_pages)
