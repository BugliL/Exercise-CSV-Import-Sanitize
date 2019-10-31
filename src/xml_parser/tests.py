import unittest

from src.xml_parser.parser import ExcelXMLParser


class XMLClassTest(unittest.TestCase):

    def test_no_errors(self):
        """When Readed throw class Than no errors"""
        x = ExcelXMLParser(filepath="src/data/table1.xml")
        header = x.header()

        header_list = [
            'ID', 'Rev', 'Description', 'Category', 'Document Code', 'Customer Code', 'Type',
            'ARI', 'Penalty', 'Purchase', 'Executor', 'Controller', 'Approver', 'Owner',
            'Note', 'Date Forecast', 'Due Date', 'Transmittal', 'Status', 'RifDow', 'Link',
        ]

        self.assertEqual(header, header_list)
