import unittest

from src.xml.XML_mask import ExcelXMLParser


class XMLClassTest(unittest.TestCase):

    # def shortDescription(self):
    #     return "Given: csv file"

    def setUp(self):
        pass

    def test_no_errors(self):
        """When Readed throw class Than no errors"""
        x = ExcelXMLParser(filepath="src/data/table1.xml")
        table = x.tables()[0]
        header_list = [
            'ID', 'Rev', 'Description', 'Category', 'Document Code', 'Customer Code', 'Type', 'ARI',
            'Penalty', 'Purchase', 'Executor', 'Controller', 'Approver', 'Owner', 'Note', 'Date Forecast',
            'Due Date', 'Transmittal', 'Status', 'RifDow', 'Link'
        ]

        self.assertEqual(table.header, header_list)
