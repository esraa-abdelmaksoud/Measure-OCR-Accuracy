import pandas as pd
import os
import re
import Levenshtein

files_path = r'/mnt/D/excel_selected'
output_path = r'/mnt/D/excel_concat_test.xlsx'


def concat_text(files_path: str, output_path: str):
    files = os.listdir(files_path)
    df_path = os.path.join(files_path, files[0])
    df = pd.read_excel(df_path, usecols=['file', 'page', 'block', 'x0', 'y0', 'x1', 'y1',
                                           'text', 'transcription'])
    # for f in range(1,len(files)):
    for f in range(1,4):
        df2_path = os.path.join(files_path, files[f])
        df2 = pd.read_excel(df2_path, usecols=['file', 'page', 'block', 'x0', 'y0', 'x1',
                                               'y1', 'text', 'transcription'])
        df = pd.concat([df, df2])
    
    # df.to_excel(output_path)
    return df


def get_all_text(df):
    df = df.reset_index()
    ocr_str, tran_str = "", ""
    for r in range(df.shape[0]):
        cur_ocr = df.loc[r,'text']
        cur_tran = df.loc[r,'transcription']
        
        # Concat text from all cells
        if pd.isna(cur_ocr) == False:
            ocr_str += (str(cur_ocr)+" ")
        if pd.isna(cur_tran) == False:
            tran_str += (str(cur_tran)+" ")
    
    # Remove line breaks and extra spaces
    ocr_str = re.sub("\n|\t|\s+", " ", ocr_str)
    tran_str = re.sub("\n|\t|\s+", " ", tran_str)

    # split words and characters
    ocr_words = ocr_str.split()
    tran_words = tran_str.split()
    ocr_chars = list(ocr_str)
    tran_chars = list(tran_str)
    return ocr_words, tran_words, ocr_chars, tran_chars

def calculate_rates(ocr_words, tran_words, ocr_chars, tran_chars):

    wer = Levenshtein.distance(tran_words, ocr_words) / len(tran_words)

    cer = Levenshtein.distance(tran_chars, ocr_chars) / len(tran_chars)
    return wer, cer

df = concat_text(files_path, output_path)
ocr_words, tran_words, ocr_chars, tran_chars = get_all_text(df)
wer, cer = calculate_rates(ocr_words, tran_words, ocr_chars, tran_chars)
# Print results
print(f"WER: {wer:.2%}")
print(f"CER: {cer:.2%}")
