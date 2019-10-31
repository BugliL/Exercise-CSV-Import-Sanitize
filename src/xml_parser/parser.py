from collections import namedtuple
from typing import Sequence

from xml.etree import ElementTree as ET

Table = namedtuple('Table', 'header')  # TypeVar('Table')


class ExcelXMLParser(object):
    HEADER_DATA_XPATH = 'std:Worksheet/std:Table/std:Row[1]//std:Data'
    NS = "urn:schemas-microsoft-com:office:spreadsheet"
    CELL_TAG = "{%s}Cell" % NS

    ns = {'std': NS}

    def __init__(self, filepath: str):
        self.filepath = filepath
        self.root = ET.parse(filepath).getroot()  # type: ET

    def header(self) -> Sequence[str]:
        return [v.text for v in self.root.findall(self.HEADER_DATA_XPATH, self.ns)]
