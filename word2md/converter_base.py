from typing import List
import math
import os
from datetime import date
from difflib import SequenceMatcher
from docx import Document
from docx.document import Document as Doc
import markdown
from lxml import etree
import chevron

class MarkdownDocument:
    def __init__(self, content, title, short_title, description) -> None:
        # These need to be set be the Word2MDConverter
        self.content = content
        self.title = title
        self.short_title = short_title
        self.description = description
        self.parent_docs = []

        # These are set by the ConverterManager
        self.is_extension = False
        self.source_file = None

        # Should not be changed
        self.attachments = self.collect_attachments(self.content)

        
    def collect_attachments(self, doc_section):
        attachments = []
        if 'graphics' in doc_section and doc_section['graphics']:
            attachments.extend(doc_section['graphics'])
        
        if 'table' in doc_section and doc_section['table']:
            for row in doc_section['table']['rows']:
                for cell in row['cells']:
                    for paragraph in cell['paragraphs']:
                        if 'graphics' in paragraph and paragraph['graphics']:
                            attachments.extend(paragraph['graphics'])

        if 'sections' in doc_section and doc_section['sections']:
            for section in doc_section['sections']:
                attachments.extend(self.collect_attachments(section))

        return attachments

class Word2MDConverter:
    CONVERTER_TYPE = 'Test Case'

    def __init__(self, document, no_emf=False):
        if isinstance(document, Doc):
            self.document = document
        else:
            self.document = Document(document)
        self.no_emf = no_emf
        self.is_extension = False

        self.mml_transform = etree.XSLT(etree.parse(os.path.join(os.path.dirname(__file__), 'xsl', 'omml2mml_v2.xsl')))
        self.remove_namespaces = etree.XSLT(etree.parse(os.path.join(os.path.dirname(__file__), 'xsl', 'remove_namespaces.xsl')))

    def convert(self) -> List[MarkdownDocument]:
        '''
        Converts the Word document into one or more outputs.
        Each ouput has the format
        {
            'path': Relative path where the output should be stored,
            'markdown': Markdown content as string
            'attachments': List with attachments (figures) 
            [ 
                {
                    'type': Type of attachment (default 'figure'),
                    'name': Name of attachment file,
                    'data': Data of attachment
                }            
            ]
        }   
        '''
        to_md_docs = self.internal_convert()

        return to_md_docs


    def internal_convert(self) -> List[MarkdownDocument]:
        '''
        Returns a dictionary with the following:
        {
            'content': The content that should be converted to Markdown,
            'path': Path where the output should be saved
        }
        '''
        raise RuntimeError(f'{self.__class__.__name__} must implement the method {self.internal_convert.__name__}.')
    
    
    def parse_table(self, table):
        raw_table = self.create_table_dict(len(table.rows), len(table.columns))
        for r, row in enumerate(table.rows):
            raw_row = self.create_row_dict()
            if r==0:
                raw_row['first'] = True
            for c, cell in enumerate(self.get_raw_cells(row)):
                raw_cell = self.create_cell_dict(self.is_heading(cell['this'], r, c), self.get_text(cell['this']), self.get_cell_contents(cell['this']), cell['colspan'])
                raw_row['cells'].append(raw_cell)
            raw_table['rows'].append(raw_row)
        return raw_table
    
    def create_table_dict(self, nr_rows=0, nr_cols=0, create_empty_cells=False):
        table = {'nr_rows': nr_rows, 'nr_cols': nr_cols, 'rows': []}
        if create_empty_cells:
            for r in range(nr_rows):
                table['rows'].append(self.create_row_dict(first=(r == 0), nr_empty_cells=nr_cols))
        return table

    
    def create_row_dict(self, first=False, nr_empty_cells=0):
        row = {'first': first, 'cells': []}
        if nr_empty_cells > 0:
            for c in range(nr_empty_cells):
                row['cells'].append(self.create_cell_dict())
        return row
    
    def create_cell_dict(self, is_heading=False, text=None, paragraphs=[], colspan=1):
        return {
            'is_heading': is_heading, 
            'text': text, 
            # 'html_text': markdown.markdown(get_text(c['this'])), 
            # 'graphics': get_inline_graphics(c['this'], document),
            # 'equations': get_inline_equations(c['this']),
            'paragraphs': paragraphs,
            'colspan': colspan
        }
    
    def create_simple_cell_dict(self, text, is_heading=False, colspan=1):
        cell = self.create_cell_dict(is_heading=is_heading, text=text, paragraphs=[self.create_simple_paragraph_dict(text)], colspan=colspan)
        return cell

    def get_raw_cells(self, row):
        cells = []
        row_cells = row.cells
        for c in row_cells:
            if len(cells) == 0 or c != cells[-1]['this']:
                cell = {'this': c, 'colspan': row_cells.count(c)}
                cells.append(cell)
        return cells

    def is_heading(self, cell, row_nr, col_nr):
        raise RuntimeError(f'{self.__class__.__name__} must implement the method {self.is_heading.__name__}.')

    def get_text(self, cell, lineseparator='\n'):
        text = lineseparator.join(self.get_paragraph_text(p) for p in cell.paragraphs)
        
        # See if there is a table within the cell that has text
        table_texts = lineseparator.join(self.get_table_text(t, lineseparator=lineseparator) for t in cell.tables)

        return (lineseparator.join(t for t in [text, table_texts] if t)).strip()

    def get_table_text(self, table, lineseparator='\n'):
        text = lineseparator.join(self.get_text(c, lineseparator=lineseparator) for c in table._cells)
        return text
    
    def get_paragraph_text(self, paragraph):
        prefix = ''
        if self.is_bullet_list(paragraph):
            level = self.get_numbering_level(paragraph)
            prefix = '    '.join(['' for _ in range(level)]) + '- '
        elif self.is_numbered_list(paragraph):
            level = self.get_numbering_level(paragraph)
            prefix = '    '.join(['' for _ in range(level)]) + '1. '
        return prefix + paragraph.text

    def is_bullet_list(self, paragraph):
        return self.get_numbering_format(paragraph) == 'bullet'

    def is_numbered_list(self, paragraph):
        fmt = self.get_numbering_format(paragraph)
        return fmt is not None and fmt != 'bullet'

    def get_numbering_level(self, paragraph):
        lvl = self.get_numbering_lvl(paragraph)
        if lvl is not None:
            try:
                lvl_indent = lvl.find('w:pPr/w:ind', namespaces=lvl.nsmap)
                if lvl_indent is not None:
                    level = math.floor(lvl_indent.left / 500000) + 1
                else:
                    level = int(self.get_value_of_attribute(lvl, 'ilvl')) + 1

                return level
            except:
                pass
        return 0

    def get_numbering_lvl(self, paragraph):
        document_part = paragraph.part
        namespaces = paragraph._element.nsmap
        w_namespace = namespaces['w']
        p_numbering = paragraph._element.find('*/w:numPr', namespaces=namespaces)
        if p_numbering is not None:
            att_val = '{' + w_namespace + '}val'
            ilvl = p_numbering.find('w:ilvl', namespaces=namespaces)
            numId = p_numbering.find('w:numId', namespaces=namespaces)
            if ilvl is not None and numId is not None:
                abstract_num_id = document_part.numbering_part.element.find(
                    'w:num[@w:numId="' + self.get_attr_val(numId) + '"]/w:abstractNumId', 
                    namespaces=namespaces)
                if abstract_num_id is not None:
                    xpath_str = ('w:abstractNum[@w:abstractNumId="' + self.get_attr_val(abstract_num_id) + '"]' + 
                        '/w:lvl[@w:ilvl="' + self.get_attr_val(ilvl) + '"]')
                    num_level = document_part.numbering_part.element.find(xpath_str, namespaces=namespaces)
                    return num_level
        return None

    def get_attr_val(self, element):
        return self.get_value_of_attribute(element, 'val')

    def get_value_of_attribute(self, element, attribute_name):
        namespaces = element.nsmap
        w_namespace = namespaces['w']
        attribute = '{' + w_namespace + '}' + attribute_name
        return element.get(attribute)

    def get_numbering_format(self, paragraph):
        num_level = self.get_numbering_lvl(paragraph)
        if num_level is not None:
            xpath_str = 'w:numFmt'
            num_fmt = num_level.find(xpath_str, namespaces=num_level.nsmap)
            if num_fmt is not None:
                return self.get_attr_val(num_fmt)
        return None    

    def get_cell_contents(self, cell, lineseparator='\n'):
        contents = []

        current_list_paragraph = []

        for paragraph in cell.paragraphs:
            p = {}

            if paragraph.text:
                if self.is_numbered_list(paragraph) or self.is_bullet_list(paragraph):
                    current_list_paragraph.append(self.get_paragraph_text(paragraph))
                    continue
                    
                p['html_text'] = self.do_markdown(self.get_paragraph_text(paragraph))
                p['raw_text'] = self.get_paragraph_text(paragraph)

            graphics = self.get_inline_graphics(paragraph, self.document)
            if graphics:
                p['graphics'] = graphics

            equations = self.get_inline_equations(paragraph)
            if equations:
                p['equations'] = equations

            if p:
                if current_list_paragraph:
                    contents.append({
                        'html_text': self.do_markdown(lineseparator.join(current_list_paragraph)),
                        'raw_text': lineseparator.join(current_list_paragraph)
                    })
                    current_list_paragraph = []
                contents.append(p)

        if current_list_paragraph:
            contents.append({
                'html_text': self.do_markdown(lineseparator.join(current_list_paragraph)),
                'raw_text': lineseparator.join(current_list_paragraph)
            })

        for t in cell.tables:
            for c in t._cells:
                contents.extend(self.get_cell_contents(c, lineseparator=lineseparator))

        return contents
    
    def create_paragraph_dict(self):
        return {
            'html_text': None,
            'raw_text': None,
            'graphics': None,
            'equations': None,
        }
    
    def create_simple_paragraph_dict(self, text):
        return {
            'html_text': self.do_markdown(text),
            'raw_text': text,
            'graphics': [],
            'equations': [],
        }

    def do_markdown(self, text):
        return markdown.markdown(text)

    def get_inline_graphics(self, word_part, document):
        try:
            element = word_part._element
        except:
            return []

        drawing_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"

        image_parts = []
        graphics = []
        for drawing in element.findall('*//w:drawing', namespaces=element.nsmap):
            namespaces = element.nsmap
            if drawing_ns not in namespaces.values():
                namespaces['a'] = drawing_ns
            blip = drawing.find('*//a:blip[@r:embed]', namespaces=namespaces)
            if blip is not None:
                graphic = {}
                graphic_id = blip.embed
                image_part = document.part.related_parts[graphic_id]         
                image_parts.append(image_part)
        
        for imagedata in element.findall('*//v:imagedata', namespaces=document._element.nsmap):
            graphic = {}
            image_id = imagedata.get('{{{0}}}id'.format(imagedata.nsmap['r']))
            image_part = document.part.related_parts[image_id]                
            image_parts.append(image_part)

        for image_part in image_parts:
            image_path = image_name = os.path.basename(image_part.partname)
            if self.no_emf:
                if '.' in image_name and image_name.endswith('.emf'):
                    image_name = '.'.join(image_name.split('.')[:-1]) + '.png'
            graphic['name'] = image_name
            graphic['path'] = image_path
            graphic['data'] = image_part._blob
            graphics.append(graphic)
            
        return graphics
    
    def save_graphic(self, text, render):
        print('Saving the graphic ' + render(text))
        return ''

    def get_inline_equations(self, word_part):
        try:
            element = word_part._element
        except:
            return []
        
        equations = []
        for equation in element.findall('*//m:oMath', namespaces=element.nsmap):
            eq = {}
            mml = self.mml_transform(equation)
            mml = self.remove_namespaces(mml.getroot())
            eq['mml'] = etree.tostring(mml, encoding='unicode')
            equations.append(eq)

        return equations    

    def is_bold(self, paragraph):
        for r in paragraph.runs:
            if not r.bold:
                return False
        return True 
    
        
    def get_cell_content_from_table(self, table_dict, text_to_find, col_delta=1, row_delta=0, compare_callable=None):
        compare_callable = compare_callable or self.strings_equal

        rows = table_dict['rows']
        for r, row in enumerate(rows):
            cells = row['cells']
            for c, cell in enumerate(cells):
                if compare_callable(cell['text'], text_to_find):
                    row_data = r + row_delta
                    cells_data = c + col_delta
                    if len(rows) > row_data and row_data >= 0 and len(rows[row_data]['cells']) > cells_data and cells_data >= 0:
                        return rows[row_data]['cells'][cells_data]['text']
        return None
        
                
    def new_section(self, heading=None, level=2, text=None, graphics=None, table=None, sub_sections=None, header=None, **extra):
        graphics = [] if graphics is None else graphics
        sub_sections = [] if sub_sections is None else sub_sections
        return {'heading': heading, 'section_level': '#'*level, 'text': text, 'graphics': graphics, 'table': table, 'sections': sub_sections, 'header': header, **extra}
    
    def create_section_header(self, title, link_title, mtime, description):
        return {
            'title': str(title).strip(),
            'link_title': str(link_title).strip(),
            'mtime': str(mtime).strip(),
            'description': str(description).strip()
        }
    
    def compare_strings(self, s1, s2):
        s = SequenceMatcher(lambda x: x in ' \t', s1.strip().lower(), s2.strip().lower())
        return s.ratio()

    def strings_equal(self, s1, s2):
        if type(s1) == str and type(s2) == str:
            return self.compare_strings(s1, s2) > 0.8
        return False
    
    def match_strings(self, string, strings_to_match):
        string_ratios = [{'string': s2, 'ratio': self.compare_strings(string, s2)} for s2 in strings_to_match]
        best_match = max(string_ratios, key=lambda x: x['ratio'])
        if self.strings_equal(string, best_match['string']):
            return best_match['string']
        return None