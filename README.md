# OCR Accuracy Measurement Tool

Measuring the OCR accuracy or the error rate requires having the ground truth. This repository includes a solution that was used to calculate the word error rate (WER) and the character error rate (CER) of OCRed PDF files.

## Methodology
This tool is in multiple steps to measure OCR (Optical Character Recognition) accuracy. It works as follows:
1) The text is OCRed using OCRmyPDF.
2) The text blocks, coordinates, block image, file name, and page name are written to an Excel file using XlsxWriter and PyMUPDF.
3) Two versions of the page images are saved. One with detected text areas covered in Green rectangles using OpenCV, and the other in original state.
4) An annotator draws red non-overlapping rectangles that cover text that wasn't detected.
5) The red rectangle areas are detected and sliced from the original images to be added also to Excel files.
6) All sliced image areas that were OCRed and that weren't OCRed are transcribed to have the ground truth.
7) All transcribed text and all OCRed text columns are concatenated.
8) The word error rate (WER) and character error rate (CER) are calculated using Levenshtein distance. 


## Installation
The tool is written in Python 3 and requires several dependencies. To install the dependencies, run the following command:
```
pip install -r requirements.txt
```

## Usage
The tool can be run from the command line using the following command:
```
python main.py input_dir output_dir
```
The input_dir is the directory containing the input images and output_dir is the directory where the output files will be saved. The input images should be in JPEG or PNG format.

The output files will be saved in the output_dir in the following format:
output.xlsx: an Excel file containing the extracted text and its coordinates.
annotated: a directory containing the annotated images with green text areas and red rectangles covering the areas where the text was not detected.
ground_truth: a directory containing the ground truth text files for each input image.
ocr_text: a directory containing the OCR-generated text files for each input image.

## License
This tool is released under the MIT License. See the LICENSE file for more details.
