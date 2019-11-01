from datetime import datetime
from collections import namedtuple
from functools import reduce
from typing import Sequence

from xml.etree import ElementTree as ET, ElementTree

Table = namedtuple('Table', 'header')  # TypeVar('Table')


class Cell(object):
    CELL_DATA_XPATH = 'std:Data'
    BASIC_NAMESPACE = 'urn:schemas-microsoft-com:office'
    ATTRIBUTE_INDEX_XPATH = '{%s}Index' % (BASIC_NAMESPACE + ':spreadsheet')
    ATTRIBUTE_TYPE_XPATH = '{%s}Type' % (BASIC_NAMESPACE + ':spreadsheet')

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
        return self.element.find(self.CELL_DATA_XPATH, self.ns).text

    @property
    def data(self):
        return self.element.find(self.CELL_DATA_XPATH, self.ns)

    @property
    def value(self):
        return self.get_function()(self.text)

    @property
    def has_index(self):
        return self.ATTRIBUTE_INDEX_XPATH in self.element.attrib

    @property
    def type(self):
        return self.data.attrib[self.ATTRIBUTE_TYPE_XPATH]

    def _get_index(self):
        return int(self.element.attrib[self.ATTRIBUTE_INDEX_XPATH])

    def get_index(self, i: int):
        return self._get_index() - 1 if self.has_index else i + 1

    def get_function(self):
        return self.FORMAT_DICT.get(self.type, str)


class ExcelXMLParser(object):
    fields_to_ignore = ['Rev', 'Due Date', 'Transmittal', 'Status', ]
    BASIC_NAMESPACE = 'urn:schemas-microsoft-com:office'

    ns = {
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

    def __init__(self, filepath: str):
        self.filepath = filepath
        self.root = ET.parse(filepath).getroot()  # type: ET

    def header(self) -> Sequence[str]:
        return [v.text for v in self.root.findall(self.HEADER_DATA_XPATH, self.ns)]

    def parse_rows(self) -> Sequence[dict]:

        header = self.header()
        lst = []
        for row in self.root.findall(self.DATA_ROWS_XPATH, self.ns)[1:]:
            i = -1
            data = {k: None for k in header if k not in self.fields_to_ignore}
            for c in row.findall(self.CELL_XPATH, self.ns):
                cell = Cell(element=c, namespaces=self.ns, headers=header)
                i = cell.get_index(i)
                k = header[i]
                if cell.data is not None and k not in self.fields_to_ignore:
                    data[k] = cell.value

            lst.append(data)

        return lst
