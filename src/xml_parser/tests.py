import unittest
from dataclasses import dataclass

from src.xml_parser.parser import ExcelXMLParser, Row


@dataclass
class DjangoModel(object):
    pk = None
    description = None
    category = None
    document_code = None
    customer_code = None
    type = None
    arip = None
    penalty = None
    purchase = None
    executor = None
    controller = None
    approver = None
    owner = None
    note = None
    date_forecast = None
    link = None
    rifdow = None


class XMLClassTest(unittest.TestCase):

    def setUp(self):
        self.parser = ExcelXMLParser(filepath="src/data/table1.xml")

    def test_header(self):
        """When Readed throw class Than get header"""
        header_list = [
            x.lower().replace(' ', '_')
            for x in ['ID', 'Rev', 'Description', 'Category', 'Document Code', 'Customer Code', 'Type',
                      'ARI', 'Penalty', 'Purchase', 'Executor', 'Controller', 'Approver', 'Owner',
                      'Note', 'Date Forecast', 'Due Date', 'Transmittal', 'Status', 'RifDow', 'Link', ]
        ]

        self.assertEqual(header_list, self.parser.header())

    def test_parsing_rows(self):
        """When readed throw class Than parse rows"""
        row = self.parser.parse_rows()[0]

        self.assertEqual(651, row['id'])
        self.assertEqual('GENERAL ARRANGEMENT DRAWING', row['description'])
        self.assertEqual('PLANT', row['category'])
        self.assertEqual('CRD700220', row['document_code'])

        row = self.parser.parse_rows()[1]
        self.assertEqual(None, row['note'])

        row = self.parser.parse_rows()[2]
        self.assertEqual('Prova', row['note'])

        row = self.parser.parse_rows()[-1]
        self.assertEqual('', row['link'])

    def test_populate_model(self):
        """given empty model when to_model called than return filled model"""

        EXCLUDE_PARAM_LIST = ['id', 'rev', 'status', 'transmittal', 'due_date', ]
        model = DjangoModel()
        row = self.parser.objects(EXCLUDE=EXCLUDE_PARAM_LIST)[0]  # type: Row
        row.fill(model)

        self.assertEqual('CRD700220', model.document_code)
        self.assertEqual('GENERAL ARRANGEMENT DRAWING', model.description)
        self.assertRaises(AttributeError, getattr, model, 'transmittal')
        self.assertRaises(AttributeError, getattr, model, 'rev')
        self.assertRaises(AttributeError, getattr, model, 'due_date')
        self.assertRaises(AttributeError, getattr, model, 'status')
