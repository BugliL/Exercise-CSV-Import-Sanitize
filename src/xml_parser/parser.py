from collections import namedtuple
from typing import Sequence

from xml.etree import ElementTree as ET

Table = namedtuple('Table', 'header')  # TypeVar('Table')


class ExcelXMLParser(object):

    def __init__(self, filepath: str):
        self.root = ET.parse(filepath).getroot()  # type: ET

    def header(self) -> Sequence[str]:
        header_list = [
            'ID', 'Rev', 'Description', 'Category', 'Document Code', 'Customer Code', 'Type', 'ARI',
            'Penalty', 'Purchase', 'Executor', 'Controller', 'Approver', 'Owner', 'Note', 'Date Forecast',
            'Due Date', 'Transmittal', 'Status', 'RifDow', 'Link'
        ]
        return header_list
