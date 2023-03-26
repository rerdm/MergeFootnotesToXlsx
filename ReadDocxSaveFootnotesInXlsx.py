import sys

from openpyxl import Workbook
from docx2python import docx2python
import re

class ReadDocFile:
    def __init__(self,word_doc_name):
        self.__word_doc_name = word_doc_name
        self.__new_txt_name = "docx_to_txt.txt"

    def doc_to_txt(self):

        document = docx2python(docx_filename=self.__word_doc_name)
        file = open(self.__new_txt_name, "w", errors="ignore")
        file.writelines(document.text)

        return self.__new_txt_name


class ReadTxtFileForFootnotes:
    def __init__(self,doc_txt_file):
        self.__doc_txt_file = doc_txt_file
        self.__xlxs_footnote_workbook = "footnotes_file.xlsx"

        self.footnote_workbook = Workbook()
        self.footnote_worksheet = self.footnote_workbook.active
        self.__line_couner = 1;

        self.footnote_worksheet.column_dimensions['A'].width = 20
        self.footnote_worksheet.column_dimensions['B'].width = 130
        self.footnote_worksheet.column_dimensions['C'].width = 130


        self.footnote_worksheet['A1'] = "Number of Footnote"
        self.footnote_worksheet['B1'] = "Footnote text"
        self.footnote_worksheet['C1'] = "Text from reference"


    def read_foot_notes(self):
        with open(self.__doc_txt_file) as file:

            for line in file:

                footnote = re.search("^(footnote)", line)
                if(footnote!=None):
                    self.__line_couner = self.__line_couner + 1
                    foot_note_id = line.split(")")[0]
                    footnote_number = foot_note_id.split("footnote")[1]
                    footnote_text = line.split(")")[1]
                    #print(footnote_number)
                    #print(footnote_text)

                    self.footnote_worksheet['A'+str(self.__line_couner)] = footnote_number
                    self.footnote_worksheet['B' + str(self.__line_couner)] = footnote_text.replace("Vgl. ","")


    def read_references(self):

        referece_found = False
        count_empty_new_lines = 0;
        line_empty = []
        new_line = ""
        counter = 1



        with open(self.__doc_txt_file) as file:

            for line in file:

                reference_section = re.search("^(Literaturverzeichnis)", line)

                if(reference_section!=None and referece_found == False):
                    referece_found = True
                    continue

                if(referece_found):

                    if line.split(" ")[0] =="<a":
                        continue

                    if (line.split(" ")[0]):

                        new_line = line.split(" ")

                        if (new_line):
                            counter = counter + 1

                            #print(counter)

                            #self.footnote_worksheet['C' + str(counter)] = new_line
                            if new_line[0] != '\n':
                                print(new_line)



        self.footnote_workbook.save(self.__xlxs_footnote_workbook)


if __name__ == '__main__':

    word_document = "Implementierung_reading_footnotes.docx"

    convert_doc_to_txt = ReadDocFile(word_document)
    txt_file_name = convert_doc_to_txt.doc_to_txt()

    read_txt_file_for_footnotes = ReadTxtFileForFootnotes(doc_txt_file=txt_file_name)
    read_txt_file_for_footnotes.read_foot_notes()

    read_txt_file_for_footnotes.read_references()








