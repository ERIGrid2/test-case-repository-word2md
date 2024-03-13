from word2md.converter_base import Word2MDConverter
from word2md.test_case import TestCaseConverter
from word2md.system_configuration import SystemConfigurationConverter
from word2md.control_functions import ControlFunctionsConverter

def is_tc_doc(document):
    for t, table in enumerate(document.tables):
        cell = table.cell(0, 0)
        if cell.text.strip().lower() == 'name of the test case':
            return True
    return False

def is_sc_doc(document):
    for t, table in enumerate(document.tables):
        cell = table.cell(0, 0)
        if cell.text.strip().lower() == 'system configuration identification':
            return True
    return False

def is_cf_document(document):
    for t, table in enumerate(document.tables):
        cell = table.cell(0, 0)
        if cell.text.strip().lower() == 'functional description':
            return True
    return False

def get_converter(document, no_emf=False) -> Word2MDConverter:
    converter = None
    if is_tc_doc(document):
        converter = TestCaseConverter(document, no_emf=no_emf)
    elif is_sc_doc(document):
        converter = SystemConfigurationConverter(document, no_emf=no_emf)
        converter.is_extension = True
    elif is_cf_document(document):
        converter = ControlFunctionsConverter(document, no_emf=no_emf)
        converter.is_extension = True
    
    return converter