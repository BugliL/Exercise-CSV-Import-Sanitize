from datetime import datetime
from collections import namedtuple
from functools import reduce
from typing import Sequence

from xml.etree import ElementTree as ET, ElementTree

Table = namedtuple('Table', 'header')  # TypeVar('Table')


class ExcelXMLParser(object):
    fields_to_ignore = ['Rev', 'Due Date', 'Transmittal', 'Status', ]

    NS = 'urn:schemas-microsoft-com:office:spreadsheet'
    ns = {
        '': NS, 'std': NS, 'ss': NS,
        'o': "urn:schemas-microsoft-com:office:office",
        'x': "urn:schemas-microsoft-com:office:excel",
        'html': "http://www.w3.org/TR/REC-html40",
    }

    ATTRIBUTE_INDEX_XPATH = '{%s}Index' % NS
    ATTRIBUTE_TYPE_XPATH = '{%s}Type' % NS

    HEADER_DATA_XPATH = 'std:Worksheet/std:Table/std:Row[1]//std:Data'
    DATA_ROWS_XPATH = 'std:Worksheet/std:Table/std:Row'
    CELL_XPATH = 'std:Cell'
    CELL_DATA_XPATH = 'std:Data'

    def __init__(self, filepath: str):
        self.filepath = filepath
        self.root = ET.parse(filepath).getroot()  # type: ET

    def header(self) -> Sequence[str]:
        return [v.text for v in self.root.findall(self.HEADER_DATA_XPATH, self.ns)]

    def parse_rows(self) -> Sequence[dict]:

        select_index = lambda c, i: \
            int(c.attrib[self.ATTRIBUTE_INDEX_XPATH]) - 1 \
                if self.ATTRIBUTE_INDEX_XPATH in c.attrib else i + 1

        fn = {
            'String': lambda x: str(x).strip(),
            'Boolean': bool,
            'Number': eval,
            'DateTime': lambda x: datetime.strptime(x, "%Y-%m-%dT%H:%M:%S.%f")
        }

        header = self.header()
        lst = []
        for row in self.root.findall(self.DATA_ROWS_XPATH, self.ns)[1:]:
            i = -1
            data = {k: None for k in header if k not in self.fields_to_ignore}
            for c in row.findall(self.CELL_XPATH, self.ns):
                i = select_index(c, i)
                d = c.find(self.CELL_DATA_XPATH, self.ns)
                k = header[i]
                if d is not None and k not in self.fields_to_ignore:
                    format = fn.get(d.attrib[self.ATTRIBUTE_TYPE_XPATH], str)
                    data[k] = format(d.text)

            lst.append(data)

        return lst
