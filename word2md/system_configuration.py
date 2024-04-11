from typing import List, Union
import re

from word2md.converter_base import Word2MDConverter
from word2md.markdown_document import MarkdownDocument, MarkdownSection, MarkdownTable
from word2md.helpers import strings_equal

class TableBasedConverter(Word2MDConverter):
    def __init__(self, document, force_png_graphics=False):
        super().__init__(document, force_png_graphics)

    def internal_convert(self) -> List[MarkdownDocument]:
        parsed_doc = self.parse_document(self.document)

        return [parsed_doc]

    def parse_document(self, document) -> MarkdownDocument:
        tables = []

        # parse tables
        for table in document.tables:
            md_table = self.parse_table(table)
            tables.append(md_table)
        
        return self.create_document_from_tables(tables)

    def create_document_from_tables(self, tables : List[MarkdownTable]) -> MarkdownDocument:
        raise RuntimeError(f'{self.__class__.__name__} must implement the method {self.create_document_from_tables.__name__}.')
    
    def get_table_heading(self, table : MarkdownTable) -> str:
        if table.rows and table.rows[0].cells:
            first_cell = table.rows[0].cells[0]
            if first_cell.is_heading:
                return first_cell.text
        return None
    
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
    
    def new_table_section(self, table : MarkdownTable, parent : Union[MarkdownDocument, MarkdownSection], heading : str = None, remove_first_row : bool = False) -> MarkdownSection:
        section_heading = heading if heading else self.get_table_heading(table)
        remove_first_row = remove_first_row if remove_first_row is not None else not heading

        section = MarkdownSection(heading=section_heading)
        
        if remove_first_row:
            table.rows.pop(0)
        section.tables.append(table)

        if isinstance(parent, MarkdownDocument):
            parent.sections.append(section)
        elif isinstance(parent, MarkdownSection):
            section.level = parent.level + 1
            parent.sub_sections.append(section)

        return section

class SystemConfigurationConverter(TableBasedConverter):
    CONVERTER_TYPE = 'System Configuration'

    def __init__(self, document, no_emf=False):
        super().__init__(document, no_emf)

    def create_document_from_tables(self, tables: List[MarkdownTable]) -> MarkdownDocument:
        md_doc = MarkdownDocument()
        component_desc_sec = None
        sc_id = None
        sc_desc = None
        for table in tables:
            heading = self.get_table_heading(table)
            if strings_equal(heading, 'Component description'):
                if component_desc_sec is None:
                    component_desc_sec = MarkdownSection(heading='Component descriptions')
                    md_doc.sections.append(component_desc_sec)
                self.new_table_section(table, component_desc_sec, heading=table.get_cell_content('Class ID', 1, 0))
                continue
            if strings_equal(heading, 'System Configuration Identification'):
                sc_id = table.get_cell_content('System configuration ID', 0, 1)
                sc_desc = table.get_cell_content('Name', 0, 1)
            self.new_table_section(table, md_doc)
        md_doc.title = 'System Configuration ' + sc_id if sc_id else 'System Configuration'
        md_doc.short_title = sc_id if sc_id else 'System Configuration'
        md_doc.description = sc_desc if sc_desc else 'This system configuration does not have a description'
        return md_doc
    