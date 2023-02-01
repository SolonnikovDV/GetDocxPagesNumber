import zipfile
import re
import glob

# local import
import config as cfg

DIR = cfg.dir


# calculate total sum odf pages in docx files
def get_pagr_count(dir: str) -> str:
    total_count = 0

    for doc_path in get_list_of_docx_files(dir):
        archive = zipfile.ZipFile(f"{doc_path}", "r")
        ms_data = archive.read("docProps/app.xml")
        archive.close()
        app_xml = ms_data.decode("utf-8")

        # filtering xml values
        regex = r"<(Pages|Slides)>(\d)</(Pages|Slides)>"

        matches = re.findall(regex, app_xml, re.MULTILINE)
        match = matches[0] if matches[0:] else [0, 0]
        # get value of sheet number in a string format
        count = match[1]
        # cast str to int
        total_count += int(count)
    return total_count


# get list of all files with extinction of 'docx'
def get_list_of_docx_files(dir: str) -> list:
    return glob.glob(f"{dir}*.docx")

print(get_list_of_docx_files(dir=DIR))
print(get_pagr_count(dir=DIR))