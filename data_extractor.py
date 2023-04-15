# If you encounter an infinate loop, it means there are not 1000 non-corrupted files
# Comment line 222 to solve it
import fitz  # pip install pymupdf
import os
import shutil
import cv2  # pip install opencv-python
import xlsxwriter  # pip install XlsxWriter
import random
import sys
import time

main_dir = r"/mnt/D/preview_reports_spatial_OCR/"

required_pages = 1000


def make_directory(main_dir: str, new_folder_name: str) -> None:
    """
    Creates a new directory in the given directory with the specified name. If a directory with the same name
    already exists, it will be replaced.

    Args:
        main_dir (str): The path of the directory where the new directory will be created.
        new_folder_name (str): The name of the new directory.

    Returns:
        None
    """
    # Create or replace directory
    dir = os.path.join(main_dir, new_folder_name)
    if os.path.isdir(dir):
        shutil.rmtree(dir)
    os.mkdir(dir)


def prepare_files(files: list) -> tuple:
    """
    Selects a random PDF file from the list of files and a random page from that file. Converts the selected page
    to an image and creates new file paths for the image, the extracted text in Excel format, and a color version of the
    image. Returns the new file paths along with the page number, the PyMuPDF document object, and the name of the PDF
    file.

    Args:
        files (list): A list of PDF file names.

    Returns:
        tuple: A tuple containing the paths to the Excel file, image file, and colored image file, the page number, the
               PyMuPDF document object, and the name of the selected PDF file.
    """
    # Select file and page
    rand_file = random.randint(0, len(files) - 1)
    file = files[rand_file]
    pdf_path = os.path.join(main_dir, files[rand_file])
    doc = fitz.open(pdf_path)
    page_num = random.randint(0, len(doc) - 1)

    # Create paths
    img_path = os.path.join(main_dir, "image", f"{file[:-4]}-{page_num}.png")
    excel_path = os.path.join(main_dir, "excel", f"{file [:-4]}-{page_num}.xlsx")
    color_path = os.path.join(main_dir, "color", f"{file [:-4]}-{page_num}.png")
    return excel_path, img_path, color_path, page_num, doc, file


def extract_data(page_num: int, img_path: str, doc: str) -> list:
    """
    Extracts text from the specified page of the given PyMuPDF document object and saves a PNG image of the page to
    the specified file path. Returns a list of dictionaries containing information about each text block on the page.
    Args:
        page_num (int): The number of the page to extract text from.
        img_path (str): The path of the file to save the PNG image of the page to.
        doc (str): The PyMuPDF document object.

    Returns:
        list: A list of dictionaries containing information about each text block on the page.
    """
    # Extract text
    page = doc.load_page(page_num)
    text_list = page.get_text("blocks")

    # Write image
    pix = page.get_pixmap()
    pix.save(img_path)
    doc.close()

    return text_list


def create_sheet(excel_path: str) -> tuple:
    """
    Creates a new Excel workbook and worksheet with the specified file path. Writes the column headers to the
    worksheet and returns both the worksheet and workbook objects.

    Args:
        excel_path (str): The path of the Excel file to create.

    Returns:
        tuple: A tuple containing the new worksheet and workbook objects.
    """
    workbook = xlsxwriter.Workbook(excel_path)
    worksheet = workbook.add_worksheet()

    cols = [
        "file",
        "page",
        "block",
        "x0",
        "y0",
        "x1",
        "y1",
        "text",
        "transcription",
        "image",
    ]
    for c in range(len(cols)):
        worksheet.write(0, c, cols[c])
    return worksheet, workbook


def write_data(
    img_path: str,
    text_list: str,
    page_num: int,
    main_dir: str,
    file: str,
    worksheet,
    color_path: str,
) -> None:
    """
    Extracts text snippets from an image, saves them to a specified directory, and writes related data to an Excel worksheet.

    Args:
        img_path (str): Path to the input image file.
        text_list (str): A list of lists containing the coordinates and text of the snippets to be extracted.
        page_num (int): The page number of the input image file.
        main_dir (str): The main directory to which the snippets will be saved.
        file (str): The filename of the input image file.
        worksheet: An openpyxl worksheet object where the extracted data will be written.
        color_path (str): Path to the output image file that will have color-coded rectangles around the extracted snippets.

    Returns:
        None.
    """
    row = 1
    img = cv2.imread(img_path)
    img_color = img.copy()

    # Going through extracted text blocks data
    for i in range(len(text_list[:-1])):
        try:

            # Save snip
            x0, y0, x1, y1 = (
                int(text_list[i][0]),
                int(text_list[i][1]),
                int(text_list[i][2]),
                int(text_list[i][3]),
            )
            snip = img[y0:y1, x0:x1]
            snip_path = os.path.join(
                main_dir, "snip", f"{file[:-4]}-{page_num}-{text_list[i][5]}.png"
            )
            cv2.imwrite(snip_path, snip)

            # Removing dirty spaces from text
            clean_text = text_list[i][4].strip()

            # Adjusting worksheet
            worksheet.set_row(row, (y1 - y0) // 2.5)

            # Writing data to worksheet
            worksheet.insert_image(
                f"J{row+1}", snip_path, {"x_scale": 0.5, "y_scale": 0.5}
            )
            values = [file, page_num, text_list[i][5], x0, y0, x1, y1, clean_text]
            for v in range(len(values)):
                worksheet.write(row, v, values[v])

            # Change area color
            img_color[y0:y1, x0:x1] = (0, 255, 0)

            row += 1

        except:
            pass
    # Write colored image
    cv2.imwrite(color_path, img_color)


def main(main_dir: str, required_pages: int) -> None:
    """
    Main function that prepares and processes files for data extraction.

    Args:
        main_dir (str): the path to the main directory.
        required_pages (int): the number of pages to process.

    Returns:
        None.
    """
    # Make required directories
    for dir_name in ["excel", "image", "color", "snip"]:
        make_directory(main_dir, dir_name)

    files = os.listdir(main_dir)

    # Process the required number of pages
    count = 0

    while count < required_pages:
        try:

            excel_path, img_path, color_path, page_num, doc, file = prepare_files(files)
            text_list = extract_data(page_num, img_path, doc)
            worksheet, workbook = create_sheet(excel_path)
            write_data(
                img_path, text_list, page_num, main_dir, file, worksheet, color_path
            )

            workbook.close()
            count += 1

            # Remove used file
            files.remove(file)
        except:
            pass
        # Print progress bar
        time.sleep(0.01)
        sys.stdout.write("\r")
        sys.stdout.write(
            "Files are being processed... {:.0f}%".format(
                (100 / (required_pages) * count)
            )
        )
        sys.stdout.flush()

    # Remove cropped snips
    shutil.rmtree(os.path.join(main_dir, "snip"))


if __name__ == "__main__":
    main(main_dir, required_pages)
