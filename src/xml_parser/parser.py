from xml.dom import minidom
from collections import namedtuple
from typing import TypeVar, Sequence


class ExcelTable(object):
    pass


Table = namedtuple('Table', 'header')  # TypeVar('Table')


class ExcelXMLParser(object):

    def __init__(self, filepath: str):
        self.document = minidom.parse(filepath)

    def tables(self) -> Sequence[Table]:
        header_list = [
            'ID', 'Rev', 'Description', 'Category', 'Document Code', 'Customer Code', 'Type', 'ARI',
            'Penalty', 'Purchase', 'Executor', 'Controller', 'Approver', 'Owner', 'Note', 'Date Forecast',
            'Due Date', 'Transmittal', 'Status', 'RifDow', 'Link'
        ]
        return [(Table(header=header_list))]
