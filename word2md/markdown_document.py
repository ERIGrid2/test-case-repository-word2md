from typing import List, Dict, Any
from dataclasses import dataclass, field
import json

import markdown

from word2md.helpers import strings_equal

class MarkdownBase:
    def encode(self):
        return vars(self)
    
    def to_dict(self):
        return json.loads(json.dumps(self, default=lambda mb: mb.encode()))
        
    def json_dumper(self, o):
        if isinstance(o, MarkdownBase):
            return o.encode()
        return o
    
@dataclass
class MarkdownDocument(MarkdownBase):
    title : str = None
    short_title : str = None
    description : str = None

    sections : List['MarkdownSection'] = field(default_factory=list)
    parent_docs : List['MarkdownDocument'] = field(default_factory=list)

    source_file : str = None
    is_extension : bool = False

    @property
    def attachments(self) -> List['MarkdownGraphic']:
        return self.collect_attachments()

    def collect_attachments(self):
        attachments = []

        for section in self.sections:
            for p in section.paragraphs:
                attachments.extend(p.graphics)

            for t in section.tables:
                for r in t.rows:
                    for c in r.cells:
                        for p in c.paragraphs:
                            attachments.extend(p.graphics)
            
        return attachments
    
    def encode(self):
        d = super().encode()
        d['attachments'] = [a.to_dict() for a in self.attachments]
        return d

@dataclass
class MarkdownContent(MarkdownBase):
    sections : List['MarkdownSection'] = field(default_factory=list)

@dataclass
class MarkdownParagraphContainer(MarkdownBase):
    paragraphs : List['MarkdownParagraph'] = field(default_factory=list)

    def add_simple_paragraph(self, text, position=-1, replace=False):
        if position >= 0 or position < len(self.paragraphs):
            if replace:
                self.paragraphs[position] = MarkdownParagraph(text=text)
            else:
                self.paragraphs.insert(position, MarkdownParagraph(text=text))
        else:
            self.paragraphs.append(MarkdownParagraph(text=text))


@dataclass
class MarkdownSection(MarkdownParagraphContainer):
    level : int = 2
    heading : str = None

    tables : List['MarkdownTable'] = field(default_factory=list)
    sub_sections : List['MarkdownSection'] = field(default_factory=list)

    @property
    def section_level(self) -> str:
        return '#' * self.level
    
    def encode(self):
        d = super().encode()
        d['section_level'] = self.section_level
        return d

@dataclass
class MarkdownParagraph(MarkdownBase):
    text : str = None
    html_text : str = None
    graphics : List['MarkdownGraphic'] = field(default_factory=list)
    equations : List['MarkdownEquation'] = field(default_factory=list)
    
    @property
    def markdown_text(self) -> str:
        if self.html_text:
            return self.html_text
        return self.do_markdown(self.text)

    def do_markdown(self, text):
        return markdown.markdown(text)
    
    def encode(self):
        d = super().encode()
        d['markdown_text'] = self.markdown_text
        return d


@dataclass
class MarkdownGraphic(MarkdownBase):
    name : str = None
    src : str = None
    data : bytes = None

    def encode(self):
        return {'name': self.name, 'src': self.src, 'data': '--'}

@dataclass
class MarkdownEquation(MarkdownBase):
    mml : str = None

@dataclass
class MarkdownTable(MarkdownBase):
    rows : List['MarkdownTableRow'] = field(default_factory=list)
    
    def add_simple_row(self, *args, row_nr=-1, heading_cols=[], replace=False):
        heading_cols = [heading_cols] if not type(heading_cols) == list else heading_cols
        row_nr = -1 if not type(row_nr) == int else row_nr

        row = MarkdownTableRow()
        for i, arg in enumerate(args):
            row.add_simple_cell(str(arg), is_heading=(i in heading_cols))

        if row_nr >= 0 and row_nr < len(self.rows):
            if replace:
                self.rows[row_nr] = row
            else:
                self.rows.insert(row_nr, row)
        else:
            self.rows.append(row)

    def get_cell_content(self, text_to_find, col_delta=1, row_delta=0, compare_callable=None):
        compare_callable = compare_callable or strings_equal

        for r, row in enumerate(self.rows):
            cells = row.cells
            for c, cell in enumerate(cells):
                if compare_callable(cell.text, text_to_find):
                    row_data = r + row_delta
                    cells_data = c + col_delta
                    if len(self.rows) > row_data and row_data >= 0 and len(self.rows[row_data].cells) > cells_data and cells_data >= 0:
                        return self.rows[row_data].cells[cells_data].text
        return None


@dataclass
class MarkdownTableRow(MarkdownBase):
    cells : List['MarkdownTableCell'] = field(default_factory=list)

    def add_simple_cell(self, text, is_heading=False, colspan=1, column=-1, replace=False):
        cell = MarkdownTableCell(is_heading=is_heading, colspan=colspan)
        cell.add_simple_paragraph(text)
        if column >= 0 and column < len(self.cells):
            if replace:
                self.cells[column] = cell
            else:
                self.cells.insert(column, cell)
        else:
            self.cells.append(cell)

@dataclass
class MarkdownTableCell(MarkdownParagraphContainer):
    is_heading : bool = False
    colspan : int = 1

    @property
    def text(self) -> str:
        return '\n'.join([p.text for p in self.paragraphs])
    
    def encode(self):
        d = super().encode()
        d['text'] = self.text
        return d
            
