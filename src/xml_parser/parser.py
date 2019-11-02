from datetime import datetime
from collections import namedtuple
from functools import reduce
from typing import Sequence, List

from xml.etree import ElementTree as ET, ElementTree


# Table = namedtuple('Table', 'header')  # TypeVar('Table')


class Contants(object):
    fields_to_ignore = ['Rev', 'Due Date', 'Transmittal', 'Status', ]
    BASIC_NAMESPACE = 'urn:schemas-microsoft-com:office'

    ALL_NS = {
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

    def __init__(self, element: ElementTree, headers: Sequence[str]):
        self.element = element
        self.headers = headers

    @property
    def text(self):
        return self.element.find(Contants.CELL_DATA_XPATH, Contants.ALL_NS).text

    @property
    def data(self):
        return self.element.find(Contants.CELL_DATA_XPATH, Contants.ALL_NS)

    @property
    def value(self):
        return self.get_function()(self.text)

    @property
    def has_index(self):
        return Contants.ATTRIBUTE_INDEX_XPATH in self.element.attrib

    @property
    def type(self):
        return self.data.attrib[Contants.ATTRIBUTE_TYPE_XPATH]

    def _get_index(self):
        return int(self.element.attrib[Contants.ATTRIBUTE_INDEX_XPATH])

    @property
    def cell_index(self):
        return self._get_index()

    def get_index(self, i: int):
        return self._get_index() - 1 if self.has_index else i + 1

    def get_function(self):
        return self.FORMAT_DICT.get(self.type, str)


class Table(object):
    def __init__(self, table_element: ElementTree):
        self.root = table_element
        self._header(table_element)
        self._body(table_element)

    def _header(self, table_element: ElementTree):
        return [
            v.text for v in \
            table_element.findall(Contants.HEADER_DATA_XPATH, Contants.ALL_NS)
        ]

    def _body(self, table_element: ElementTree):
        raise NotImplementedError()


class ExcelXMLParser(object):

    def __init__(self, filepath: str):
        self.filepath = filepath
        self.root = ET.parse(filepath).getroot()  # type: ET

    def header(self) -> Sequence[str]:
        return [v.text for v in self.root.findall(Contants.HEADER_DATA_XPATH, Contants.ALL_NS)]

    def parse_rows(self) -> Sequence[dict]:
        header = self.header()
        lst = []
        for row_element in self.root.findall(Contants.DATA_ROWS_XPATH, Contants.ALL_NS)[1:]:
            row = row_element
            i = -1
            data = {k: None for k in header}
            for c_element in row.findall(Contants.CELL_XPATH, Contants.ALL_NS):
                cell = Cell(element=c_element, headers=header)
                i = cell.get_index(i)
                k = header[i]
                if cell.data is not None:
                    data[k] = cell.value

            lst.append(data)

        return lst
