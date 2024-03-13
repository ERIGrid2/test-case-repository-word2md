import re
import os
import argparse
from datetime import date

from word2md.converter_base import Word2MDConverter, MarkdownDocument

class TestCaseConverter(Word2MDConverter):

    def __init__(self, document, no_emf=False):
        super().__init__(document, no_emf)        
        self.test_case_headline_regex = re.compile('Test Case\s(.*)')
        self.test_specification_headline_regex = re.compile('Test Specification\s(.*)')
        self.experiment_specification_headline_regex = re.compile('Experiment Specification\s(.*)')

        self.test_case = {'content': None, 'title': None, 'short_title': None, 'description': None, 'id': None}
        self.test_specs = []
        self.exp_specs = []

    def internal_convert(self):
        test_specifications = self.find_test_specifications()
        experiment_specifications = self.find_experiment_specifications()

        tc_md_doc = None
        ts_md_docs = {}

        parsed_documents = []

        # parse docx file
        test_spec_tables = []
        exp_spec_tables = []
        for table in self.document.tables:
            if self.is_test_case(table):
                test_case = self.parse_test_case(table)
                tc_md_doc = MarkdownDocument(test_case, test_case['header']['title'], test_case['header']['link_title'], test_case['header']['description'])
                parsed_documents.append(tc_md_doc)
            elif self.is_test_specification(table):
                test_spec_tables.append(table)
            elif self.is_experiment_specification(table):
                exp_spec_tables.append(table)

        if tc_md_doc is not None:
            for number_test_specs, table in enumerate(test_spec_tables):
                test_spec = test_specifications[number_test_specs] if number_test_specs < len(test_specifications) else {}
                content = self.parse_test_specification(table, test_spec)
                
                md_doc = MarkdownDocument(content, content['header']['title'], content['header']['link_title'], content['header']['description'])
                md_doc.parent_docs = [tc_md_doc]
                ts_md_docs[content['id']] = md_doc

                parsed_documents.append(md_doc)

            for number_experiment_specs, table in enumerate(exp_spec_tables):
                exp_spec = experiment_specifications[number_experiment_specs] if number_experiment_specs < len(experiment_specifications) else {}
                content = self.parse_experiment_specification(table, exp_spec)

                md_doc = MarkdownDocument(content, content['header']['title'], content['header']['link_title'], content['header']['description'])
                md_doc.parent_docs = [tc_md_doc]
                if content['test_spec_id'] in ts_md_docs:
                    md_doc.parent_docs.append(ts_md_docs[content['test_spec_id']])

                parsed_documents.append(md_doc)

        return parsed_documents

    def is_test_case(self, table):
        cell = table.cell(0, 0)
        if cell.text.strip().lower() == 'name of the test case':
            return True
        return False

    def is_test_specification(self, table):
        cell = table.cell(1, 0)
        if cell.text.strip().lower() == 'title of test':
            return True
        return False

    def is_experiment_specification(self, table):
        cell = table.cell(1, 0)
        if cell.text.strip().lower() == 'title of experiment':
            return True
        return False

    def is_test_case_headline(self, paragraph):
        text = paragraph.text
        if self.test_case_headline_regex.match(text):
            if self.is_bold(paragraph):
                return True
        return False

    def is_test_specification_headline(self, paragraph):
        text = paragraph.text
        if self.test_specification_headline_regex.match(text):
            if self.is_bold(paragraph):
                return True
        return False

    def is_experiment_specification_headline(self, paragraph):
        text = paragraph.text
        if self.experiment_specification_headline_regex.match(text):
            if self.is_bold(paragraph):
                return True
        return False

    def is_qualification_strategy_headline(self, paragraph):
        text = paragraph.text
        if text.strip() == 'Qualification Strategy':
            if self.is_bold(paragraph):
                return True
        return False

    def is_mapping_headline(self, paragraph):
        text = paragraph.text
        if text.strip() == 'Mapping to Research Infrastructure':
            if self.is_bold(paragraph):
                return True
        return False
    
    def is_heading(self, cell, row_nr, col_nr):
        return False
    
    def add_simple_row_heading(self, table_dict, *args, row_nr=-1):
        return self.add_simple_row(table_dict, *args, row_nr=row_nr, heading_cols=[0])

    def add_simple_row(self, table_dict, *args, row_nr=-1, heading_cols=[]):
        heading_cols = [heading_cols] if not type(heading_cols) == list else heading_cols
        row_nr = -1 if not type(row_nr) == int else row_nr

        row = self.create_row_dict()
        for i, arg in enumerate(args):
            row['cells'].append(self.create_simple_cell_dict(str(arg), is_heading=(i in heading_cols)))
        
        if not table_dict['rows'] or row_nr == 0:
            row['first'] = True
            for r in table_dict['rows']:
                r['first'] = False
        table_dict['rows'].insert(row_nr, row)       
        

    def parse_test_case(self, table):
        sections = []
        tc_id = ''

        id_section = self.new_section(heading='Identification')
        id_table = self.create_table_dict(nr_rows=5, nr_cols=2)
        id_section['table'] = id_table
        sections.append(id_section)

        re_author_version = re.compile('Author:?\s+(.*)\s+Version:?\s+(.*)')
        re_project_date = re.compile('Project:?\s+(.*)\s+Date:?\s+(.*)')
        is_qs = False
        for p in self.document.paragraphs:
            text = p.text
            if self.is_test_case_headline(p):
                tc_id = self.test_case_headline_regex.match(text).group(1).strip()
                self.add_simple_row_heading(id_table, 'ID', tc_id)
            if re_author_version.match(text):
                self.add_simple_row_heading(id_table, 'Author', re_author_version.match(text).group(1).strip())
                self.add_simple_row_heading(id_table, 'Version', re_author_version.match(text).group(2).strip())
            if re_project_date.match(text):
                self.add_simple_row_heading(id_table, 'Project', re_project_date.match(text).group(1).strip())
                self.add_simple_row_heading(id_table, 'Date', re_project_date.match(text).group(2).strip())
            if self.is_test_specification_headline(p):
                break
            if self.is_qualification_strategy_headline(p):
                is_qs = True
                qs_section = self.new_section(heading='Qualification Strategy', level=2, text='')
            elif is_qs:                
                qs_section['text'] = (qs_section['text'] + '\n' + self.get_paragraph_text(p)).strip()
                graphics = self.get_inline_graphics(p, self.document)
                if len(graphics) > 0:
                    qs_section['graphics'].extend(graphics)

        tc_table = self.parse_tc_table(table)

        tc_desc = self.get_table_content_from_heading(tc_table, 'Name of the Test Case')

        tc_table_section = self.new_section(heading='Test Case Definition')
        tc_table_section['table'] = tc_table
        sections.append(tc_table_section)
        
        sections.append(qs_section)

        tc_header = self.create_section_header(
            f'Test Case {tc_id}', 
            tc_id or 'Test Case', 
            date.today().isoformat(), 
            tc_desc or 'A Test Case'
        )

        tc_content = self.new_section(sub_sections=sections, header=tc_header)

        return tc_content
    
    def get_table_content_from_heading(self, table_dict, heading, compare_callable=None):
        return self.get_cell_content_from_table(table_dict, heading, 1, 0)

    def find_test_specifications(self):
        test_specs = []
        
        is_mapping = False
        for p in self.document.paragraphs:
            text = p.text
            if self.is_test_specification_headline(p):
                test_spec = {}
                test_spec['ID'] = {'desc': self.test_specification_headline_regex.match(text).group(1).strip()}
                test_specs.append(test_spec)
            
            if self.is_experiment_specification_headline(p):
                break
            if self.is_mapping_headline(p):
                is_mapping = True
                test_spec = test_specs[-1]
                test_spec['Mapping to Research Infrastructure'] = {'desc': '', 'graphics': []}
            elif is_mapping:
                test_spec_mapping = test_specs[-1]['Mapping to Research Infrastructure']
                test_spec_mapping['desc'] = test_spec_mapping['desc'] + '\n' + self.get_paragraph_text(p) if test_spec_mapping['desc'] else self.get_paragraph_text(p)
                graphics = self.get_inline_graphics(p, self.document)
                if len(graphics) > 0:
                    test_spec_mapping['graphics'].extend(graphics)
        return test_specs

    def find_experiment_specifications(self):
        experiment_specs = []
        
        for p in self.document.paragraphs:
            text = p.text
            if self.is_experiment_specification_headline(p):
                exp_spec = {}
                exp_spec['ID'] = {'desc': self.experiment_specification_headline_regex.match(text).group(1).strip()}
                experiment_specs.append(exp_spec)
        return experiment_specs

    def parse_test_specification(self, table, test_spec):
        ts_table_section = self.new_section('Test Specification Definition')
        ts_table = self.parse_tc_table(table)
        ts_id = ''
        for key in test_spec.keys():
            if key == 'ID':
                ts_id = test_spec[key]['desc']
                self.add_simple_row_heading(ts_table, key, ts_id, row_nr=0)
            else:
                self.add_simple_row_heading(ts_table, key, test_spec[key]['desc'])

        ts_table_section['table'] = ts_table

        ts_desc = self.get_table_content_from_heading(ts_table, 'Title of Test')
        
        extra = {
            'id': ts_id,
            'path': ts_id
        }
        
        ts_header = self.create_section_header(
            f'Test Specification {ts_id}', 
            ts_id or 'Test Specification', 
            date.today().isoformat(), 
            ts_desc or 'A Test Specification'
        )

        ts_content = self.new_section(sub_sections=[ts_table_section], header=ts_header, **extra)  
        return ts_content

    def parse_experiment_specification(self, table, experiment_spec):
        es_tabl_section = self.new_section('Test Specification Definition')
        es_table = self.parse_tc_table(table)
        for key in experiment_spec.keys():
            if key == 'ID':
                es_id = experiment_spec[key]['desc']
                self.add_simple_row_heading(es_table, key, es_id, row_nr=0)
            else:
                self.add_simple_row_heading(es_table, key, experiment_spec[key]['desc'])
        es_tabl_section['table'] = es_table
        
        es_desc = self.get_table_content_from_heading(es_table, 'Title of Experiment')        
        es_header = self.create_section_header(
            f'Experiment Specification {es_id}', 
            es_id or 'Experiment Specification', 
            date.today().isoformat(), 
            es_desc or 'An Experiment Specification'
        )

        extra = {
            'id': es_id,
            'test_spec_id': self.get_table_content_from_heading(es_table, 'Reference to Test Specification'),
            'path': os.path.join(self.get_table_content_from_heading(es_table, 'Reference to Test Specification'), es_id)
        }

        es_content = self.new_section(sub_sections=[es_tabl_section], header=es_header, **extra)  
        return es_content
    
    def parse_tc_table(self, table):
        tc_table = self.parse_table(table)
        for row in tc_table['rows']:
            for index, first_cell in enumerate(row['cells']):
                if first_cell['text']:
                    break
            id = first_cell['text'].split('\n')[0]
            id = id.split(':')[0].strip()
            row['cells'][index] = self.create_simple_cell_dict(id, is_heading=True, colspan=first_cell['colspan'])
        return tc_table
    

def get_test_cases(folder_or_doc, recurse=False):
    files_to_convert = []

    if os.path.isdir(folder_or_doc):
        for f in os.scandir(folder_or_doc):
            if f.is_file() and f.path.endswith('.docx'):
                files_to_convert.append(f.path)
            elif recurse and f.is_dir():
                files_to_convert.extend(get_test_cases(f, recurse=recurse))
    elif os.path.isfile(folder_or_doc) and folder_or_doc.endswith('.docx'):
        files_to_convert.append(folder_or_doc)
    
    return files_to_convert

if __name__ == '__main__':
    excel_template_default = os.path.join(os.path.dirname(__file__), 'template', 'HTD_TEMPLATE_V1.3.xlsx')

    parser = argparse.ArgumentParser(description='Converts test cases according to the ERIGrid HTD Template from Word into Excel files.')
    parser.add_argument('path', help='Path to either a Word file or a folder. If a folder is provided, all Word files in that folder will be converted.')
    parser.add_argument('-t', '--excel-template', help='Path to the Excel template that should be used. Standard: {0}'.format(excel_template_default),
                        default=os.path.join(os.path.dirname(os.path.abspath(__file__)), excel_template_default))
    parser.add_argument('-f', '--create-folder', help='Saves the Excel file and extracted images to a folder with the name of Word file.', 
                        action='store_true')
    parser.add_argument('-c', '--copy-word-file', help='Copies the Word file into the new folder', action='store_true')
    parser.add_argument('-r', '--recurse', help='Recurse subfolders', action='store_true')
    parser.add_argument('-u', '--update', help='Updates an already existing Excel file to a new template version and adds any new content. Can be used to update previously converted test cases.', action='store_true')
    args = parser.parse_args()    

    doc_filename = args.path
    template_path = args.excel_template
    create_folder = args.create_folder
    copy_word_file = args.copy_word_file
    recurse = args.recurse
    update = args.update

    files_to_convert = get_test_cases(doc_filename, recurse=recurse)
        
    for f in files_to_convert:
        print('\nConverting {0}'.format(f))
        tcc = TestCaseConverter()
        tcc.convert(f, template_path, create_folder=create_folder, copy_word_file=copy_word_file, update=update)