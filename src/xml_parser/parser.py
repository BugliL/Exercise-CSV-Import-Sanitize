from datetime import datetime
from collections import namedtuple
from functools import reduce
from typing import Sequence

from xml.etree import ElementTree as ET, ElementTree

Table = namedtuple('Table', 'header')  # TypeVar('Table')


class Namespaces(object):
    fields_to_ignore = ['Rev', 'Due Date', 'Transmittal', 'Status', ]
    BASIC_NAMESPACE = 'urn:schemas-microsoft-com:office'

    ALL = {
        '': BASIC_NAMESPACE + ':spreadsheet',
        'std': BASIC_NAMESPACE + ':spreadsheet',
        'ss': BASIC_NAMESPACE + ':spreadsheet',
        'o': BASIC_NAMESPACE + ":office",
        'x': BASIC_NAMESPACE + ":excel",
        'html': "http://www.w3.org/TR/REC-html40",
    }

    HEADER_DATA_XPATH = 'std:Worksheet/std:Table/std:Row[1]//std:Data'
    DATA_ROWS_XPATH = 'std:Worksheet/std:Table/std:Row'
    CELL_XPATH = 'std:Cell'
    CELL_DATA_XPATH = 'std:Data'
    BASIC_NAMESPACE = 'urn:schemas-microsoft-com:office'
    ATTRIBUTE_INDEX_XPATH = '{%s}Index' % (BASIC_NAMESPACE + ':spreadsheet')
    ATTRIBUTE_TYPE_XPATH = '{%s}Type' % (BASIC_NAMESPACE + ':spreadsheet')


class Cell(object):
    FORMAT_DICT = {
        'String': lambda x: str(x).strip(),
        'Boolean': bool,
        'Number': eval,
        'DateTime': lambda x: datetime.strptime(x, "%Y-%m-%dT%H:%M:%S.%f")
    }

    def __init__(self, element: ElementTree, namespaces: dict, headers: Sequence[str]):
        self.element = element
        self.ns = namespaces
        self.headers = headers

    @property
    def text(self):
        return self.element.find(Namespaces.CELL_DATA_XPATH, Namespaces.ALL).text

    @property
    def data(self):
        return self.element.find(Namespaces.CELL_DATA_XPATH, Namespaces.ALL)

    @property
    def value(self):
        return self.get_function()(self.text)

    @property
    def has_index(self):
        return Namespaces.ATTRIBUTE_INDEX_XPATH in self.element.attrib

    @property
    def type(self):
        return self.data.attrib[Namespaces.ATTRIBUTE_TYPE_XPATH]

    def _get_index(self):
        return int(self.element.attrib[Namespaces.ATTRIBUTE_INDEX_XPATH])

    def get_index(self, i: int):
        return self._get_index() - 1 if self.has_index else i + 1

    def get_function(self):
        return self.FORMAT_DICT.get(self.type, str)


class ExcelXMLParser(object):

    def __init__(self, filepath: str):
        self.filepath = filepath
        self.root = ET.parse(filepath).getroot()  # type: ET

    def header(self) -> Sequence[str]:
        return [v.text for v in self.root.findall(Namespaces.HEADER_DATA_XPATH, Namespaces.ALL)]

    def parse_rows(self) -> Sequence[dict]:
        header = self.header()
        lst = []
        for row_element in self.root.findall(Namespaces.DATA_ROWS_XPATH, Namespaces.ALL)[1:]:
            row = row_element
            i = -1
            data = {k: None for k in header}
            for c_element in row.findall(Namespaces.CELL_XPATH, Namespaces.ALL):
                cell = Cell(element=c_element, namespaces=Namespaces.ALL, headers=header)
                i = cell.get_index(i)
                k = header[i]
                if cell.data is not None:
                    data[k] = cell.value

            lst.append(data)

        return lst
