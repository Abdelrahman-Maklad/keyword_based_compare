import concurrent.futures
import os.path
import re
from multiprocessing import freeze_support
import pandas as pd
import pdfplumber
from itertools import repeat
import fitz
from tqdm import tqdm
from rich.console import Console
from rich.markdown import Markdown
from datetime import date

def sce_part_number(link: str):
    link = link.strip()
    try:
        doc = pdfplumber.open(link)
        lines_list = doc.pages[0].extract_text_lines()

        if lines_list[1]['text'].lower() == 'specifications':
            font_size = lines_list[1]['chars'][0]['size']

        else:
            return ''

        for line in lines_list[2:]:
            if int(line['chars'][0]['size']) == int(font_size):
                return line['text']
    except Exception as E:
        print(E)
        return ''


def read_and_group(excel_file_path):
    data_frame = pd.read_excel(excel_file_path)
    grouped_data = data_frame.groupby("VENDOR_CODE")
    return grouped_data


def count_of_pages(link):
    doc = fitz.open(link)
    return doc.page_count


def extract_all_text(link):
    doc = pdfplumber.open(link)
    text = "\n".join([page.extract_text_simple() for page in doc.pages])
    return text.lower().strip()


def search_for_words(keywords, text):
    matches_by_word = {keyword.strip(): [] for keyword in keywords}
    for keyword in keywords:
        matches = []
        if keyword.startswith("_"):
            matches = [text.split('\n')[int(keyword.removeprefix("_")) - 1]]
        else:
            for line in text.split("\n"):
                if keyword.strip().lower() in line.strip().lower():
                    matches.append(line.strip())

        matches = [''] if matches == [] else matches
        matches_by_word[keyword.strip()].extend(matches)
    return matches_by_word


def compare(temp_list_1, temp_list_2):
    match_list_1 = []
    match_list_2 = []
    for match in temp_list_1:
        match_list_1.append(match.replace(" ", ""))
    for match in temp_list_2:
        match_list_2.append(match.replace(" ", ""))
    match_list_1.sort()
    match_list_2.sort()
    if match_list_1 == match_list_2:
        return True
    elif match_list_1 != match_list_2:
        return False


def compare_and_mapping(matches_by_word1, matches_by_word2, part_1, part_2, part_status, keywords, keyword_mapping):
    df_list = []
    changed_features = []
    mapped_features = []

    for keyword in keywords:

        matches_list_1 = matches_by_word1[keyword.strip()]
        matches_list_2 = matches_by_word2[keyword.strip()]

        if not compare(matches_list_1, matches_list_2):
            changed_features.append(keyword.strip())

        for changed_feature in changed_features:
            for keyword_mapping_feature in keyword_mapping:
                if changed_feature.lower().strip() == keyword_mapping_feature[0].lower().strip():
                    mapped_features.append(keyword_mapping_feature[-1].lower().strip())

                    break
            else:
                mapped_features.append(changed_feature.lower().strip())

    mapped_features = list(set(mapped_features))
    mapped_features.sort()

    # start insertion in the result row
    df_list.insert(0, ", ".join(mapped_features))
    df_list.insert(0, "Yes" if changed_features != [] else "No")
    if df_list[0] == 'Yes' and part_status == "Equal":
        df_list.insert(0, "Equal Parts")
        df_list.insert(0, "")
        df_list.insert(0, "Done")
        df_list.insert(0, part_2)
        df_list.insert(0, part_1)
        df_list.insert(0, "Done")
    elif df_list[0] == 'No' and part_status == "Equal":
        df_list.insert(0, "Equal Manual")
        df_list.insert(0, "")
        df_list.insert(0, "Done")
        df_list.insert(0, part_2)
        df_list.insert(0, part_1)
        df_list.insert(0, "Done")
    else:
        df_list = []
        df_list.insert(0, "Not Compatible")
        df_list.insert(0, "Reject")

    return df_list


def compare_parts(links, text1, text2, part_keyword, vendor_code):
    part_keyword = part_keyword[0].strip().lower()
    if part_keyword != 'xxx':

        # compare part action
        for line in text1.split('\n'):
            if part_keyword in line:
                part_line_1 = line
                break

        else:
            part_line_1 = ''

        for line in text2.split('\n'):
            if part_keyword in line:
                part_line_2 = line
                break
        else:
            part_line_2 = ''

    else:
        if vendor_code == 'SCE':
            part_line_1 = sce_part_number(links[0])
            part_line_2 = sce_part_number(links[-1])

    part_1 = (re.sub(fr"(.*{part_keyword.lower()}|[0-9]+/[0-9]+/20[0-9]+|:|www.mill-max.com)", "", part_line_1)
              .removeprefix(".").replace("  ", " "))
    part_2 = (re.sub(fr"(.*{part_keyword.lower()}|[0-9]+/[0-9]+/20[0-9]+|:|www.mill-max.com)", "", part_line_2)
              .removeprefix(".").replace("  ", " "))
    if part_1.replace(" ", "").strip() == part_2.replace(" ", "").strip() and part_1.strip() != '':
        return part_1.strip(), part_2.strip(), 'Equal'
    else:
        return '_', '_', 'Not Equal'


def magic_main(links, keywords, vendor_code, part_keyword, keyword_mapping):
    final_row = []
    try:
        '''Get  number of pages in the document'''
        num_pages_1 = count_of_pages(links[0])
        num_pages_2 = count_of_pages(links[-1])

        if num_pages_2 > 50 or num_pages_1 > 50:
            return [links[0], links[-1], "".join(links).replace('\\\\10.199.104.160\\pdfs', ""), vendor_code,
                    "Reject", "Page limit exceeded"]

        text1 = extract_all_text(links[0])
        text2 = extract_all_text(links[-1])

        part_1, part_2, part_status = compare_parts(links,text1, text2, part_keyword, vendor_code)

        if part_status != 'Equal':
            final_row.insert(0, "Not Compatible")
            final_row.insert(0, "Reject")
            final_row.insert(0, vendor_code)
            final_row.insert(0, "".join(links).replace('\\\\10.199.104.160\\pdfs', ""))
            final_row.insert(0, links[-1])
            final_row.insert(0, links[0])
            return final_row

        matches_by_word1 = search_for_words(keywords, text1)
        matches_by_word2 = search_for_words(keywords, text2)

        final_row = compare_and_mapping(matches_by_word1, matches_by_word2, part_1, part_2, part_status, keywords, keyword_mapping)
        final_row.insert(1, str(len(text2)))
        final_row.insert(1, str(len(text1)))
        final_row.insert(0, vendor_code)
        final_row.insert(0, "".join(links).replace('\\\\10.199.104.160\\pdfs', ""))
        final_row.insert(0, links[-1])
        final_row.insert(0, links[0])

        return final_row

    except Exception as E:
        return [links[0], links[-1], "".join(links).replace('\\\\10.199.104.160\\pdfs', ""), vendor_code, "Error"]


'''----------------------------------------------------------------------------------------------------'''
if __name__ == '__main__':
    freeze_support()

    console = Console()

    title = '''# MAGIC SPECIAL
    > this tool is designed for the following purposes:
        1- extracting all lines with the keywords given in the keyword file "keywords.xlsx".
        2- compare between the lines and specify the changes.
        3- maps the keywords to the it's related feature in the database given in mapping file "mapping.xlsx".
    > in the "input.xlsx" input file all input are given and are grouped by the vendor code given.
    > the keywords and mapped features are grouped based on the vendor code to avoid any intersections.
    > the output is provided in separate files each named after the related supplier.
    ** each supplier must have keywords and mapping database.
        
    '''
    my_copyright = '''# © abdelrahman.maklad@siliconexpert.com'''
    title = Markdown(title)
    my_copyright = Markdown(my_copyright)
    console.print(title)
    console.print(my_copyright)

    title = '''▶PRESS ENTER TO START.'''
    title = Markdown(title)
    console.print(title)
    input()

    # start program
    with open('fixed_path.txt', 'r') as path_file:
        path = path_file.read().strip()
    # start program

    grouped_links = read_and_group('input.xlsx')

    grouped_keywords = read_and_group(f"{path}\\keywords.xlsx")

    grouped_mapping = read_and_group(f'{path}\\mapping.xlsx')

    grouped_part_numbers = read_and_group('PartNumber.xlsx')

    vendor_keywords_dict = {}
    vendor_links_dict = {}
    keyword_mapping_dict = {}
    vendor_part_numbers_dict = {}

    for vendor_code, group in grouped_links:
        vendor_links_dict[vendor_code] = list(group[['DOCUMENT', 'LATEST']].values)

    for vendor_code, group in grouped_mapping:
        keyword_mapping_dict[vendor_code] = list(group[['KEYWORD', 'MAPPING']].values)

    for vendor_code, group in grouped_keywords:
        vendor_keywords_dict[vendor_code] = list(group['KEYWORD'])

    for vendor_code, group in grouped_part_numbers:
        vendor_part_numbers_dict[vendor_code] = list(group['PART_NUMBER'])

    '''-------------------------------------------------------------------------------------------------'''
    # create header
    header = []
    header.extend(
        ['OLD', 'LATEST', 'OLD_LATEST', 'VENDOR_CODE', 'STATUS','CHAR_LEN_1', 'CHAR_LEN_2', 'PART 1', 'PART 2',
         'COMPARE_SATUS', 'COMMENT','PART STATUS', 'DATA_CHANGED', 'CHANGED_FEATURES'])

    output_file_name = f'magic_special_output_{str(date.today())}.txt'
    if not os.path.exists(output_file_name):
        with open(output_file_name, 'w') as output_file:
            output_file.write("\t".join(header))
            output_file.write('\n')

    # start program for every vendor
    for key_vendor_code in list(vendor_links_dict.keys()):
        try:
            one_time_count = 250
            links_list = vendor_links_dict[key_vendor_code]

            total_iterations = len(links_list)

            with tqdm(total=total_iterations, desc=f"Processing {key_vendor_code}".upper(), unit="row",
                      ncols=100) as progress_bar:

                with concurrent.futures.ProcessPoolExecutor(max_workers=5) as executor:
                    # pass links to function
                    for i in range(0, len(links_list), one_time_count):
                        batch_links = links_list[i:i + one_time_count]
                        results = executor.map(magic_main, batch_links, repeat(vendor_keywords_dict[key_vendor_code]),
                                               repeat(key_vendor_code),
                                               repeat(vendor_part_numbers_dict[key_vendor_code])
                                               , repeat(keyword_mapping_dict[key_vendor_code]))
                        # write results
                        results_list = []
                        for result in results:
                            try:
                                with open(output_file_name, 'a') as output_file:
                                    output_file.write("\t".join(result))
                                    output_file.write('\n')
                                progress_bar.update(1)
                            except Exception as E:
                                progress_bar.update(1)
                                pass
                progress_bar.set_description(f"done {key_vendor_code}".upper())
        except Exception as E:
            pass
