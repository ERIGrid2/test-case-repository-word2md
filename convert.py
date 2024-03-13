from typing import List
import os
import argparse
from docx import Document
from datetime import date
import chevron
import yaml
import logging

from word2md.converter_factory import get_converter
from word2md.converter_base import MarkdownDocument

class OutputDocument:
    def __init__(self, header=None, content=None, output_dir=None, file_name=None, attachments=None) -> None:
        self.header = header
        self.content = content
        self.output_dir = output_dir
        self.file_name = file_name
        self.attachments = attachments or []

class ConverterManager:

    def __init__(self, input_path, destination, create_folder=False, recurse=False, no_emf=False) -> None:
        self.input_path = input_path
        self.output_dir = destination
        self.create_folder = create_folder
        self.recurse = recurse
        self.no_emf = no_emf

    def to_md(self, output_doc : OutputDocument):    
        md_result = {'header': None, 'content': None}

        md_result['header'] = yaml.dump(output_doc.header)

        md_result['content'] = self.render_mustache(output_doc.content, 'MDContent.mustache')

        return self.render_mustache(md_result, 'MDDocument.mustache')
        
    def render_mustache(self, md_content, template_name):
        md_content['openbrace'] = '{'
        md_content['closebrace'] = '}'
        md_content['newline'] = '\n'
        templates_path = os.path.join(os.path.dirname(__file__), 'word2md', 'mustache')
        with open(os.path.join(templates_path, template_name), 'r') as template:
            md_file_content = chevron.render(template=template, data=md_content, partials_path=templates_path)
            return md_file_content

    def convert_file(self, doc_filename):
        document = None
        try:
            document = Document(doc_filename)
        except:
            logging.error('ERROR: Could not open Word file: {0}'.format(doc_filename))
            return
        
        converter = get_converter(document, no_emf=self.no_emf)
        if converter is None:
            logging.error('ERROR: No converter avilable for this type of document.')
            return
        
        logging.info(f'{doc_filename} -> {converter.CONVERTER_TYPE}')
        
        md_documents = converter.convert()

        for md_doc in md_documents:
            md_doc.source_file = doc_filename
            if converter.is_extension:
                md_doc.is_extension = True            

        return md_documents
            

    def make_sure_exisits(self, folder_path):
        try:
            os.makedirs(folder_path)
        except FileExistsError as e:
            pass

    def has_subfolders(self, folder_path):
        for root, directories, files in os.walk(folder_path):
            if directories:
                return True
        return False


    def get_files_to_convert(self, folder_or_doc, output_dir, recurse=False, folder_prefix=None):
        if folder_prefix is None:
            folder_prefix = folder_or_doc

        files_to_convert = []

        if os.path.isdir(folder_or_doc):
            for f in os.scandir(folder_or_doc):
                if f.is_file() and f.path.endswith('.docx'):
                    output_file_dirname = os.path.join(output_dir, os.path.relpath(os.path.dirname(f.path), folder_prefix))
                    files_to_convert.append({
                        'file_path': f.path,
                        'output_dir': output_file_dirname
                    })
                elif recurse and f.is_dir():
                    files_to_convert.extend(self.get_files_to_convert(f, output_dir, recurse=recurse, folder_prefix=folder_prefix))
        elif os.path.isfile(folder_or_doc) and folder_or_doc.endswith('.docx'):
            files_to_convert.append({
                'file_path': folder_or_doc,
                'output_dir': output_dir
            })
        
        return files_to_convert
    
    def convert(self):
        output_docs = []
        if os.path.isdir(self.input_path):
            output_docs = self.convert_folder(self.input_path, self.output_dir, recurse=self.recurse)
        elif os.path.isfile(self.input_path) and self.input_path.endswith('.docx'):
            output_docs = self.convert_files([self.input_path], self.output_dir)

        for output_doc in output_docs:
            md_str = self.to_md(output_doc)

            self.make_sure_exisits(output_doc.output_dir)

            with open(os.path.join(output_doc.output_dir, output_doc.file_name), 'w', encoding='utf-8') as output:
                output.write(md_str)

            # Print attachments
            for attachment in output_doc.attachments:
                attachment_path = os.path.join(output_doc.output_dir, attachment['path'])        
                with open(attachment_path, 'wb') as fs:
                    fs.write(attachment['data'])


    def convert_folder(self, folder, base_output_dir, recurse=False, folder_prefix=None) -> List[OutputDocument]:
        output_docs = []

        if folder_prefix is None:
            folder_prefix = folder

        if os.path.isdir(folder):
            files_to_convert = []
            output_dir = os.path.join(base_output_dir, os.path.relpath(folder, folder_prefix))
            for f in os.scandir(folder):
                if f.is_file() and f.path.endswith('.docx'):
                    files_to_convert.append(f.path)
                elif recurse and f.is_dir():
                    output_docs.extend(self.convert_folder(f, base_output_dir, recurse=recurse, folder_prefix=folder_prefix))

            output_docs.extend(self.convert_files(files_to_convert, output_dir))
        
        return output_docs

    def convert_files(self, files_to_convert, output_dir) -> List[OutputDocument]:
        output_docs = []

        md_documents : List[MarkdownDocument] = []
        for f in files_to_convert:
            md_documents.extend(self.convert_file(f))

        for md_doc in md_documents:
            md_header = {}
            md_header['title'] = self.escape_quotes(md_doc.title)
            md_header['linkTitle'] = self.escape_quotes(md_doc.short_title)
            md_header['date'] = self.escape_quotes(date.today().isoformat())
            md_header['description'] = self.escape_quotes(md_doc.description)
            md_header['weight'] = 1
            if md_doc.is_extension:
                md_header['weight'] = 10                
            
            output_file_dir = self.get_output_file_dir(output_dir, md_doc)
            file_name = '_index.md'
            
            output_docs.append(OutputDocument(header=md_header, content=md_doc.content, output_dir=output_file_dir, file_name=file_name, attachments=md_doc.attachments))

        return output_docs
    
    def escape_quotes(self, text):
        text = str(text)
        text = text.replace('\'', '\'\'')
        return text
    
    def get_output_file_dir(self, base_output_dir, md_doc : MarkdownDocument):
        folder_name = ''
        if md_doc.source_file and self.create_folder:
            folder_name = '.'.join(os.path.basename(md_doc.source_file).split('.')[:-1])

        if md_doc.is_extension or md_doc.parent_docs:
            folder_name = os.path.join(folder_name, md_doc.short_title)

        if not md_doc.parent_docs:
            return os.path.join(base_output_dir, folder_name)
        
        return os.path.join(self.get_output_file_dir(base_output_dir, md_doc.parent_docs[-1]), folder_name)



if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Converts test cases according to the ERIGrid HTD Template from Word to Markdown files.')
    parser.add_argument('path', help='Path to either a Word file or a folder. If a folder is provided, all Word files in that folder will be converted.')
    parser.add_argument('destination', help='Path to a folder where the output will be saved. If "create-folder" is true, the output folder is created.')
    parser.add_argument('-f', '--create-folder', help='Saves the Markdown file and extracted images to a folder in "destination" with the name of the Word file.', 
                        action='store_true')
    parser.add_argument('-r', '--recurse', help='Recurse subfolders', action='store_true')
    parser.add_argument('-e', '--no-emf', help='Forces graphics with file ending ".emf" to ".png".', action='store_true')
    args = parser.parse_args()    

    logging.basicConfig(format='%(asctime)s - %(message)s', level=logging.INFO)

    converter_manager = ConverterManager(args.path, args.destination, create_folder=args.create_folder, recurse=args.recurse, no_emf=args.no_emf)

    logging.info(f'Conversion started for {args.path}')
    converter_manager.convert()
    logging.info(f'Conversion completed. Output files written to {args.destination}')