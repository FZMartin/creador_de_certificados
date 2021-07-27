from classes.classes import *

def main():
    app = App(dest_path='path\\to\\folder',
              word_template_path='path\\to\\docx_template',
              csv_path='path\\to\\csv',
              keywords_map=['${nombre}', '${dni}'])
    
    app.create_certificates()
    app.export_to_pdf()
    

if __name__ == '__main__':
    main()