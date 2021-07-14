import os
from docx import Document
import csv
from datetime import datetime

import docx


class App(object):
    def __init__(self, dest_path:str, word_template_path:str, csv_path:str, keywords_map:list) -> None:
        """
        Clase entorno para aplicacion de creacion de certificados
        
        :param dest_path: ubicacion de la carpeta de destino para guardar los certificados
        :param word_template: ubicacion del archivo word a utilizar como plantilla
        :param csv_path: ubicacion del archivo csv de donde se extraera informacion para certificados
        """
        self.destPath = dest_path
        self.template_path = word_template_path
        self.csv1_path = csv_path
        self.keywordsMap = keywords_map
        
        self.template = WordManage(word_template_path, self.keywordsMap)
        
        
        
    def make_certificates(self):
        """
        Metodo principal para iniciar creacion de certificados.
        """
        pass


class WordManage(Document):
    def __init__(self, word_path:str, keywords:list, ) -> None:
        """
        Clase para manipulacion rapida de archivos word
        
        :param word_path: ubicacion del archivo word 
        :param keywords: lista con las palabras claves a remplazar
        """
        self.path = word_path
        self.keywords = keywords
        self.docx = Document(self.path)
        
    def replace(self, info:list):
        """
        Metodo para reemplazo de palabras claves por valores.
        
        :param info: lista con la informacion que se reemplazara
        """
        if len(info) == len(self.keywords):
            for keyword, item in zip(self.keywords, info):
                try:
                    for table in self.docx.tables:
                        for row in table.rows:
                            for cell in row:
                                for p in cell.paragraphs:
                                    for run in p.runs:
                                        if run.text == keyword:
                                            run.text = item
                        else:
                            print(f'No se encontro la palabra clave {keyword}')
                except:
                    print(f'Error al remplazar informacion') 
        else:
            print('Deben haber iguales cantidades de argumentos y de palabras claves')
    
    def save_as(self, _dest_path:str, file_name:str):
        try:
            self.docx.save(f'{_dest_path}\\{file_name.upper()}')
        except:
            print(f'Error al guardar')


class PersonToCertificate:
    def __init__(self, _app:App, _nombre:str, _apellido:str, _dni:str):
        self.app = _app
        self.nombre = _nombre.upper()
        self.apellido = _apellido.upper()
        self.dni = self.fix(_dni)
        
    def fix(self, _dni:str):
        dni = _dni
        if '.' not in dni:
            dni = f'{dni[:2]}.{dni[2:5]}.{dni[5:]}'
            return dni
        else:
            dni.replace('.', '')
            self.fix(dni)
    
    def get_full_name(self) -> str:
        return f'{self.nombre} {self.apellido}'
    
    def make_my_certificate(self):
        pass
        
            
        
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
        elif self.csvFile.mode == 'a':
            self.writer = csv.DictWriter(self.csvFile, fields)
        else:
            print('mode must be "r" (read), "w" (write) or "a" (append)')

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
    
    def append(self, ):
        pass

    def close(self):
        """
        Funcion para cerrar archivo leido.

        :return:
        """
        self.csvFile.close()

