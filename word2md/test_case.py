import re
from word2md.converter_base import Word2MDConverter
from word2md.markdown_document import (
    MarkdownDocument,
    MarkdownSection,
    MarkdownTable
)

class TestCaseConverter(Word2MDConverter):

    def __init__(self, document, no_emf=False):
        super().__init__(document, no_emf)        
        self.test_case_headline_regex = re.compile('\s*Test\s+Case\s+(.*)')
        self.test_specification_headline_regex = re.compile('\s*Test\s+Specification\s+(.*)')
        self.experiment_specification_headline_regex = re.compile('\s*Experiment\s+Specification\s+(.*)')

        self.test_case = {'content': None, 'title': None, 'short_title': None, 'description': None, 'id': None}
        self.test_specs = []
        self.exp_specs = []

        self.tc_md_doc = None
        self.ts_md_docs = {}

    def internal_convert(self):
        test_specifications = self.find_test_specifications()
        experiment_specifications = self.find_experiment_specifications()

        parsed_documents = []

        # parse docx file
        test_spec_tables = []
        exp_spec_tables = []
        for table in self.document.tables:
            if self.is_test_case(table):
                self.tc_md_doc = self.parse_test_case(table)
                parsed_documents.append(self.tc_md_doc)
            elif self.is_test_specification(table):
                test_spec_tables.append(table)
            elif self.is_experiment_specification(table):
                exp_spec_tables.append(table)

        if self.tc_md_doc is not None:
            for number_test_specs, table in enumerate(test_spec_tables):
                test_spec = test_specifications[number_test_specs] if number_test_specs < len(test_specifications) else {}
                md_doc = self.parse_test_specification(table, test_spec)
                parsed_documents.append(md_doc)
            for number_experiment_specs, table in enumerate(exp_spec_tables):
                exp_spec = experiment_specifications[number_experiment_specs] if number_experiment_specs < len(experiment_specifications) else {}
                md_doc = self.parse_experiment_specification(table, exp_spec)
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
    
    def add_simple_row_heading(self, table : MarkdownTable, *args, row_nr=-1):
        table.add_simple_row(*args, row_nr=row_nr, heading_cols=[0])  
        
    def parse_test_case(self, table):
        test_case = MarkdownDocument()
        sections = test_case.sections
        tc_id = ''

        id_section = MarkdownSection(heading='Identification')
        id_table = MarkdownTable()
        id_section.tables.append(id_table)
        sections.append(id_section)

        qs_section_paragraphs = []

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
                qs_section_paragraphs = []
            elif is_qs:     
                qs_section_paragraphs.append(p)      

        tc_table = self.parse_tc_table(table)

        tc_desc = self.get_table_content_from_heading(tc_table, 'Name of the Test Case')

        tc_table_section = MarkdownSection(heading='Test Case Definition')
        tc_table_section.tables.append(tc_table)
        sections.append(tc_table_section)
        
        qs_section = MarkdownSection(heading='Qualification Strategy', level=2, paragraphs=self.get_content_from_paragraphs(qs_section_paragraphs))
        sections.append(qs_section)

        test_case.title = f'Test Case {tc_id}'
        test_case.short_title = tc_id or 'Test Case'
        test_case.description = tc_desc or 'A Test Case'

        return test_case
    
    def get_table_content_from_heading(self, table : MarkdownTable, heading, compare_callable=None):
        return table.get_cell_content(heading, 1, 0)

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
        test_spec_doc = MarkdownDocument()
        ts_table_section = MarkdownSection(heading='Test Specification Definition')
        test_spec_doc.sections.append(ts_table_section)
        ts_table = self.parse_tc_table(table)
        ts_table_section.tables.append(ts_table)

        ts_id = ''
        if 'ID' in test_spec:
            ts_id = test_spec['ID']['desc']
            self.add_simple_row_heading(ts_table, 'ID', ts_id, row_nr=0)

        ts_desc = self.get_table_content_from_heading(ts_table, 'Title of Test')

        test_spec_doc.title = f'Test Specification {ts_id}'
        test_spec_doc.short_title = ts_id or 'Test Specification'
        test_spec_doc.description = ts_desc or 'A Test Specification'
        
        test_spec_doc.parent_docs.append(self.tc_md_doc)
        self.ts_md_docs[test_spec_doc.short_title.strip()] = test_spec_doc

        return test_spec_doc

    def parse_experiment_specification(self, table, experiment_spec):
        es_spec_doc = MarkdownDocument()
        es_tabl_section = MarkdownSection(heading='Experiment Specification Definition')
        es_spec_doc.sections.append(es_tabl_section)
        es_table = self.parse_tc_table(table)
        es_tabl_section.tables.append(es_table)

        es_id =''
        if 'ID' in experiment_spec:
            es_id = experiment_spec['ID']['desc']
            self.add_simple_row_heading(es_table, 'ID', es_id, row_nr=0)
        
        es_desc = self.get_table_content_from_heading(es_table, 'Title of Experiment')        
        
        es_spec_doc.title = f'Experiment Specification {es_id}'
        es_spec_doc.short_title = es_id or 'Experiment Specification' 
        es_spec_doc.description = es_desc or 'An Experiment Specification'

        es_spec_doc.parent_docs.append(self.tc_md_doc)
        test_spec_id = self.get_table_content_from_heading(es_table, 'Reference to Test Specification').strip()
        if test_spec_id in self.ts_md_docs:
            es_spec_doc.parent_docs.append(self.ts_md_docs[test_spec_id])

        return es_spec_doc
    
    def parse_tc_table(self, table):
        tc_table = self.parse_table(table)
        for row in tc_table.rows:
            for first_cell in row.cells:
                if first_cell.text.strip():
                    break
            id = first_cell.paragraphs[0].text
            id = id.split(':')[0].strip()
            first_cell.paragraphs[0].text = id
            first_cell.paragraphs = first_cell.paragraphs[:1]
            first_cell.is_heading = True

        return tc_table
    