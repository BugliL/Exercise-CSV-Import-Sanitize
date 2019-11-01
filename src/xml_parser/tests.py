import datetime
import unittest
from functools import partial

from src.xml_parser.parser import ExcelXMLParser


class XMLClassTest(unittest.TestCase):

    def setUp(self):
        self.parser = ExcelXMLParser(filepath="src/data/table1.xml")

    def test_header(self):
        """When Readed throw class Than get header"""
        header_list = [
            'ID', 'Rev', 'Description', 'Category', 'Document Code', 'Customer Code', 'Type',
            'ARI', 'Penalty', 'Purchase', 'Executor', 'Controller', 'Approver', 'Owner',
            'Note', 'Date Forecast', 'Due Date', 'Transmittal', 'Status', 'RifDow', 'Link',
        ]

        self.assertEqual(self.parser.header(), header_list)

    def test_parsing_rows(self):
        """When readed throw class Than parse rows"""
        row = self.parser.parse_rows()[0]

        self.assertEqual(651, row['ID'])
        self.assertEqual('GENERAL ARRANGEMENT DRAWING', row['Description'])
        self.assertEqual('PLANT', row['Category'])

        row = self.parser.parse_rows()[1]
        self.assertEqual(None, row['Note'])

        row = self.parser.parse_rows()[2]
        self.assertEqual('Prova', row['Note'])

        row = self.parser.parse_rows()[-1]
        self.assertEqual('', row['Link'])
