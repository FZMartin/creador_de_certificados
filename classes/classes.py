import os
import docx
import csv


class App(object):
    def __init__(self, dest_path:str, word_template_path:str, csv_path:str, keywords_map:list) -> None:
        self.destPath = dest_path
        self.csv1 = csv_path
        self.template = word_template_path


class WordManage(docx):
    def __init__(self, app: App, word_path:str, keywords:list, ) -> None:
        super().__init__()
        self.path = word_path
        self.keywords = keywords


class CsvLoader(object):
    def __init__(self, csvPath, mode, fields):
        """
        Punto de inicio para la clase CSVLoader.

        :param csvPath: str path del archivo
        :param mode: str r(read)/w(write)/a(append)
        PRECAUCION: w mode sobreescribira un archivo como vacio
        :param fields:
        """
        self.csvFile = open(csvPath, mode=mode, encoding='UTF-8', newline='')
        self.fields = fields
        if self.csvFile.mode == 'r':
            self.reader = csv.DictReader(self.csvFile, fields)
        elif self.csvFile.mode == 'w':
            self.writer = csv.DictWriter(self.csvFile, fields)
        else:
            print('mode must be "r" (read) or "w" (write)')

    def getContentAsList(self, firstLineHeaders=False):
        """
        En caso de que el modo sea r, se crea un reader, con el cual se puede obtener todo el contenido del archivo como
        una lista.

        :param firstLineHeaders: bool
        :return: list
        """
        try:
            contentList = [[i[self.fields[0]],
                            i[self.fields[1]],
                            i[self.fields[2]],
                            i[self.fields[3]],
                            i[self.fields[4]]] for i in self.reader]
            if firstLineHeaders:
                return contentList[1:]
            else:
                return contentList
        except AttributeError as e:
            print(f'there is no reader.\nERROR: {e}')
            return None

    def close(self):
        """
        Funcion para cerrar archivo leido.

        :return:
        """
        self.csvFile.close()
