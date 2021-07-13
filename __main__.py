from classes.classes import *

def main():
    app = App(dest_path='',
              word_template_path='',
              csv_path='',
              keywords_map=[])
    
    app.create_certificate()

if __name__ == '__main__':
    main()