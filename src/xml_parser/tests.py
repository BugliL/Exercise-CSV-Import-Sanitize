import datetime
import unittest
from functools import partial

from src.xml_parser.parser import ExcelXMLParser


class XMLClassTest(unittest.TestCase):

    def test_header(self):
        """When Readed throw class Than get header"""
        x = ExcelXMLParser(filepath="src/data/table1.xml")
        header_list = [
            'ID', 'Rev', 'Description', 'Category', 'Document Code', 'Customer Code', 'Type',
            'ARI', 'Penalty', 'Purchase', 'Executor', 'Controller', 'Approver', 'Owner',
            'Note', 'Date Forecast', 'Due Date', 'Transmittal', 'Status', 'RifDow', 'Link',
        ]

        self.assertEqual(x.header(), header_list)

    def test_parsing_rows(self):
        """When readed throw class Than """
        x = ExcelXMLParser(filepath="src/data/table1.xml")
        row = x.parse_rows()[0]

        self.assertEqual(651, row['ID'])
        self.assertEqual('GENERAL ARRANGEMENT DRAWING', row['Description'])
        self.assertEqual('PLANT', row['Category'])

        row = x.parse_rows()[1]
        self.assertEqual(datetime.datetime(year=2019, month=2, day=7), row['Due Date'])
        self.assertEqual(None, row['Note'])

        row = x.parse_rows()[2]
        self.assertEqual('Prova', row['Note'])
