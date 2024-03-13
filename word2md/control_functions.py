from datetime import date
from word2md.system_configuration import TableBasedConverter

class ControlFunctionsConverter(TableBasedConverter):
    CONVERTER_TYPE = 'Control Functions'

    def __init__(self, document, no_emf=False):
        super().__init__(document, no_emf)

        # special sections
        self.inputs_sec = self.new_section('Inputs')
        self.outputs_sec = self.new_section('Outputs')
        self.use_cases_sec = self.new_section('Use Cases')

    def format_section(self, section, parent_section, formated_doc):
        best_match = self.match_strings(section['heading'], ['Control Function Input', 'Control Function Output', 'Use Case Example', 'Control Function Identification', 'Algorithms'])
        if best_match == 'Control Function Input':
            if self.inputs_sec not in parent_section['sections']:
                parent_section['sections'].append(self.inputs_sec)
            parent_section = self.inputs_sec
            section['section_level'] = '#' * 3

            new_heading = self.get_cell_content_from_table(section['table'], 'Name', 1, 0)
            if new_heading:
                section['heading'] = new_heading

        elif best_match == 'Control Function Output':
            if self.outputs_sec not in parent_section['sections']:
                parent_section['sections'].append(self.outputs_sec)
            parent_section = self.outputs_sec
            section['section_level'] = '#' * 3

            new_heading = self.get_cell_content_from_table(section['table'], 'Name', 1, 0)
            if new_heading:
                section['heading'] = new_heading

        elif best_match == 'Use Case Example':
            if self.use_cases_sec not in parent_section['sections']:
                parent_section['sections'].append(self.use_cases_sec)
            parent_section = self.use_cases_sec
            section['section_level'] = '#' * 3

            new_heading = self.get_cell_content_from_table(section['table'], 'Use Case Example', 1, 0)
            if new_heading:
                section['heading'] = new_heading        
        elif best_match == 'Control Function Identification':
            cf_id = self.get_cell_content_from_table(section['table'], 'Control Function ID', 0, 1)
            if cf_id:
                formated_doc['cf_id'] = cf_id
            cf_desc = self.get_cell_content_from_table(section['table'], 'Name', 0, 1)
            if cf_desc:
                formated_doc['cf_desc'] = cf_desc
        elif best_match == 'Algorithms':
            all_cells = [cell for row in section['table']['rows'] for cell in row['cells']]
            for cell in all_cells:
                if not cell['is_heading']:
                    code_block = '\n'.join(['    ' + line for line in cell['text'].split('\n')])
                    cell['paragraphs'] = [self.create_simple_paragraph_dict(code_block)]

        super().format_section(section, parent_section, formated_doc)

    def get_header(self, formated_doc):
        title = 'Control Function ' + formated_doc['cf_id'] if 'cf_id' in formated_doc else 'Control Function'
        link_title = formated_doc['cf_id'] if 'cf_id' in formated_doc else 'Control Function'
        description = formated_doc['cf_desc'] if 'cf_desc' in formated_doc else 'This control function does not have a description'
        mtime = date.today().isoformat()

        return self.create_section_header(title, link_title, mtime, description)
    
    def get_path(self, formatted_doc):
        if 'cf_id' in formatted_doc:
            return f'CtrlFun{formatted_doc["cf_id"]}' if 'cf_id' in formatted_doc else 'CtrlFun'
        return super().get_path(formatted_doc)