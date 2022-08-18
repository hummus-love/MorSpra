import os
from collections import Counter
import pprint as pp
import pandas as pd
import nltk
from nltk.tokenize import word_tokenize
import xml.etree.ElementTree as ET

#####

# MW = Moralwort
# MWs = Moralwörter

#####

german_MWs_path = r'path/to/german_Morallexikon.xlsx'
english_MWs_path = r'path/to/english_Morallexikon.xlsx'
french_MWs_path = r'path/to/french_Morallexikon.xlsx'

# read in MWs files with ALL sheets as dict
df_ger_mw = pd.read_excel(german_MWs_path, sheet_name=None, header=None)
df_eng_mw = pd.read_excel(english_MWs_path, sheet_name=None, header=None)
df_fre_mw = pd.read_excel(french_MWs_path, sheet_name=None, header=None)


def count_mws_in_texts(language, text_dir):
    if language not in {'eng', 'ger', 'fre'}:
        raise ValueError("language must be 'eng' or 'ger' or 'fre'!")

    writer = pd.ExcelWriter(path='MW_counts_txt.xlsx', engine='xlsxwriter', engine_kwargs={'options':{'strings_to_formulas': False}})

    dir = os.fsencode(text_dir)

    for file in os.listdir(dir):
        filename = os.fsdecode(file)
        file_path = os.path.join(text_dir, filename)

        # Sometimes you have to change the encoding type (utf-8, utf-16, utf-16le, ...)
        with open(file_path, 'r', encoding='utf-8') as f:
            file_contents = f.read()
        
        all_tokens = word_tokenize(file_contents)
        all_lowered_tokens = [x.lower() for x in all_tokens]
        
        print('\n', filename, len(all_lowered_tokens))
        #print(len(lowered_tokens_wo_stopwords))

        df = pd.DataFrame(columns=['Morallexikon', 'Absolute MW Counts', 'MW Counts per 10.000 Words'])

        # German
        if language == 'ger':
            lang_dict = df_ger_mw
        elif language == 'eng':
            lang_dict = df_eng_mw
        elif language == 'fre':
            lang_dict = df_fre_mw

        for i, (sheetname, sheetdata) in enumerate(lang_dict.items()):
            mw_sum = 0
            mws = sheetdata[0].to_list()
            mws_lowered = [x.lower() for x in mws]

            for token in all_lowered_tokens:
                if token in mws_lowered:
                    mw_sum += 1
            print(sheetname, mw_sum)

            relative_count = mw_sum / (len(all_lowered_tokens) / 10000)
            print("MWs per 10.000 words", relative_count)

            df.loc[i] = [sheetname, mw_sum, relative_count]

        print(df)
        df.to_excel(writer, sheet_name=filename[:30], index=False)
    
    writer.save()

    return None



myPath = r''
count_mws_in_texts('ger', myPath)
