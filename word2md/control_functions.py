from typing import List
from .markdown_document import MarkdownTable
from word2md.markdown_document import MarkdownDocument, MarkdownSection, MarkdownParagraph
from word2md.system_configuration import TableBasedConverter
from word2md.helpers import match_strings

class ControlFunctionsConverter(TableBasedConverter):
    CONVERTER_TYPE = 'Control Functions'

    def __init__(self, document, no_emf=False):
        super().__init__(document, no_emf)

    def create_document_from_tables(self, tables: List[MarkdownTable]) -> MarkdownDocument:
        md_doc = MarkdownDocument()
        inputs_sec = None
        outputs_sec = None
        use_cases_sec = None
        cf_id = None
        cf_desc = None
        for table in tables:
            heading = self.get_table_heading(table)
            best_match = match_strings(heading, ['Control Function Input', 'Control Function Output', 'Use Case Example', 'Control Function Identification', 'Algorithms'])
            if best_match == 'Control Function Input':
                if inputs_sec is None:
                    inputs_sec = MarkdownSection(heading='Inputs')
                    md_doc.sections.append(inputs_sec)
                self.new_table_section(table, inputs_sec, heading=table.get_cell_content('Name', 1, 0))
                continue
            elif best_match == 'Control Function Output':
                if outputs_sec is None:
                    outputs_sec = MarkdownSection(heading='Outputs')
                    md_doc.sections.append(outputs_sec)
                self.new_table_section(table, outputs_sec, heading=table.get_cell_content('Name', 1, 0))
                continue
            elif best_match == 'Use Case Example':
                if use_cases_sec is None:
                    use_cases_sec = MarkdownSection(heading='Use Cases')
                    md_doc.sections.append(use_cases_sec)
                self.new_table_section(table, use_cases_sec, heading=table.get_cell_content('Use Case Example', 1, 0))
                continue
            elif best_match == 'Control Function Identification':
                cf_id = table.get_cell_content('Control Function ID', 0, 1)
                cf_desc = table.get_cell_content('Name', 0, 1)
            elif best_match == 'Algorithms':
                for row in table.rows:
                    for cell in row.cells:
                        if not cell.is_heading:
                            code_block = '\n```\n' + cell.text + '\n```\n\n'
                            cell.paragraphs = [MarkdownParagraph(text=code_block, html_text=code_block)]
            self.new_table_section(table, md_doc)
        md_doc.title = 'Control Function ' + cf_id if cf_id else 'Control Function'
        md_doc.short_title = cf_id if cf_id else 'Control Function'
        md_doc.description = cf_desc if cf_desc else 'This control function does not have a description'
        return md_doc
