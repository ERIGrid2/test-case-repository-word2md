import re
from datetime import date
import uuid

from word2md.converter_base import Word2MDConverter, MarkdownDocument

class TableBasedConverter(Word2MDConverter):
    def __init__(self, document, force_png_graphics=False):
        super().__init__(document, force_png_graphics)

    def internal_convert(self) -> list:
        parsed_doc = self.parse_document(self.document)

        formatted_doc = self.format_document(parsed_doc)

        header = self.get_header(formatted_doc)
        md_doc = MarkdownDocument(formatted_doc, header['title'], header['link_title'], header['description'])

        return [md_doc]

    def parse_document(self, document):
        parsed_doc = {'tables': [], 'objects': {}}

        # parse docx file
        for t, table in enumerate(document.tables):
            raw_table = self.parse_table(table)

            table_obj = {'heading': self.get_table_heading(table), '_id_': str(uuid.uuid4()), 'raw_data': raw_table, 'sub_sections': []}

            parsed_doc['tables'].append(table_obj)              
        
        return parsed_doc
    
    def get_table_heading(self, table):
        # use the text from the first cell in the table as heading
        cells = self.get_raw_cells(table.rows[0])
        if cells and self.has_background_color(cells[0]['this']):
            return self.get_text(cells[0]['this'])
        return None
    
    def format_document(self, parsed_doc):
        formated_doc = self.new_section()

        for table in parsed_doc['tables']:
            parent_section = formated_doc

            heading = table['heading'] or 'New Section'
            section = self.new_section(heading=heading, level=2, table=table['raw_data'])

            self.format_section(section, parent_section, formated_doc)

        return formated_doc
    
    def format_section(self, section, parent_section, formated_doc):
        parent_section['sections'].append(section)
    
    def is_heading(self, cell, row_nr, col_nr):
        return self.has_background_color(cell)
    
    def has_background_color(self, cell):
        pattern = re.compile('w:fill=\"(\S*)\"')
        match = pattern.search(cell._tc.xml)
        if match:
            result = match.group(1)
            if result and result != 'auto':
                return True
        return False

class SystemConfigurationConverter(TableBasedConverter):
    CONVERTER_TYPE = 'System Configuration'

    def __init__(self, document, no_emf=False):
        super().__init__(document, no_emf)

        # special sections
        self.component_desc_sec = self.new_section('Component descriptions')

    def format_section(self, section, parent_section, formated_doc):
        if self.strings_equal(section['heading'], 'Component description'):
            if self.component_desc_sec not in parent_section['sections']:
                parent_section['sections'].append(self.component_desc_sec)
            
            parent_section = self.component_desc_sec
            section['section_level'] = '#' * 3

            new_heading = self.get_cell_content_from_table(section['table'], 'Class ID', 1, 0)
            if new_heading:
                section['heading'] = new_heading
  
        elif self.strings_equal(section['heading'], 'System Configuration Identification'):
            sc_id = self.get_cell_content_from_table(section['table'], 'System configuration ID', 0, 1)
            if sc_id:
                formated_doc['sc_id'] = sc_id
            sc_desc = self.get_cell_content_from_table(section['table'], 'Name', 0, 1)
            if sc_desc:
                formated_doc['sc_desc'] = sc_desc
          
        super().format_section(section, parent_section, formated_doc)
    
    def get_header(self, formated_doc):
        title = 'System Configuration ' + formated_doc['sc_id'] if 'sc_id' in formated_doc else 'System Configuration'
        link_title = formated_doc['sc_id'] if 'sc_id' in formated_doc else 'System Configuration'
        description = formated_doc['sc_desc'] if 'sc_desc' in formated_doc else 'This system configuration does not have a description'
        mtime = date.today().isoformat()

        return self.create_section_header(title, link_title, mtime, description)
    
    def get_path(self, formatted_doc):
        if 'sc_id' in formatted_doc:
            return f'SysConf{formatted_doc["sc_id"]}' if 'sc_id' in formatted_doc else 'SysConf'
        return super().get_path(formatted_doc)
    