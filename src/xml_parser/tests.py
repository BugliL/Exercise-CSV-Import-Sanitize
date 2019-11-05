import datetime
import unittest
from dataclasses import dataclass

from src.xml_parser.parser import ExcelXMLParser, Row


@dataclass
class DjangoModel(object):
    pk: int = None
    description: str = None
    category: str = None
    document_code: str = None
    customer_code: str = None
    type: str = None
    arip: str = None
    penalty: bool = None
    purchase: bool = None
    executor: str = None
    controller: str = None
    approver: str = None
    owner: str = None
    note: str = None
    date_forecast: datetime.datetime = None
    link: str = None
    rifdow: str = None
    project: int = None


class XMLClassTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.maxDiff = None

    def setUp(self):
        self.parser1 = ExcelXMLParser(filepath="src/data/table1.xml")
        self.parser2 = ExcelXMLParser(filepath="src/data/table2.xml")

    def test_header(self):
        """When Readed throw class Than get header"""
        header_list = [
            x.lower().replace(' ', '_')
            for x in ['ID', 'Rev', 'Description', 'Category', 'Document Code', 'Customer Code', 'Type',
                      'ARI', 'Penalty', 'Purchase', 'Executor', 'Controller', 'Approver', 'Owner',
                      'Note', 'Date Forecast', 'Due Date', 'Transmittal', 'Status', 'RifDow', 'Link', ]
        ]

        self.assertEqual(header_list, self.parser1.header())

    def test_parsing_rows(self):
        """When readed throw class Than parse rows"""
        row = self.parser1.parse_rows()[0]

        self.assertEqual(None, row['id'])
        self.assertEqual('GENERAL ARRANGEMENT DRAWING', row['description'])
        self.assertEqual('PLANT', row['category'])
        self.assertEqual('CRD700220', row['document_code'])

        row = self.parser1.parse_rows()[1]
        self.assertEqual(None, row['note'])

        row = self.parser1.parse_rows()[2]
        self.assertEqual('Prova', row['note'])

        row = self.parser1.parse_rows()[-1]
        self.assertEqual('', row['link'])

    def test_populate_model(self):
        """given empty model when to_model called than return filled model"""

        EXCLUDE_PARAM_LIST = ['id', 'rev', 'status', 'transmittal', 'due_date', ]
        model = DjangoModel()
        row = self.parser1.objects(EXCLUDE=EXCLUDE_PARAM_LIST)[0]  # type: Row
        row.fill(model)

        self.assertEqual('CRD700220', model.document_code)
        self.assertEqual('GENERAL ARRANGEMENT DRAWING', model.description)
        self.assertRaises(AttributeError, getattr, model, 'transmittal')
        self.assertRaises(AttributeError, getattr, model, 'rev')
        self.assertRaises(AttributeError, getattr, model, 'due_date')
        self.assertRaises(AttributeError, getattr, model, 'status')

    def test_functional(self):
        """Testing functional programming"""
        EXCLUDE_PARAM_LIST = ['id', 'rev', 'status', 'transmittal', 'due_date', ]

        model0 = DjangoModel()
        model0.project = 12
        row0 = self.parser1.objects(EXCLUDE=EXCLUDE_PARAM_LIST)[0]  # type: Row
        row0.fill(model0)

        model1 = DjangoModel()
        row1 = self.parser1.objects(EXCLUDE=EXCLUDE_PARAM_LIST)[1]  # type: Row
        row1.fill(model1)
        model1.project = 12
        dummy_list = [model0, model1]

        set_project = lambda x: setattr(x, 'project', 12) or x
        get_model = lambda row: DjangoModel()
        fill_model = lambda row, model: row.fill(model)

        record_list = [
            fill_model(row, set_project(get_model(row)))
            for row in self.parser1.objects(EXCLUDE=EXCLUDE_PARAM_LIST)
        ]

        self.assertListEqual(dummy_list, record_list[:2])
        self.assertEqual(dummy_list[0].description, record_list[0].description)
        self.assertEqual(dummy_list[0].document_code, record_list[0].document_code)
        self.assertEqual('CRD700220', dummy_list[0].document_code)
        self.assertEqual(12, dummy_list[0].project)
        self.assertNotEqual('ayeye', dummy_list[0].document_code)

    def test_parser2(self):
        """When Readed throw class Than remove empty rows"""
        rows = self.parser2.parse_rows()
        self.assertTrue(all([bool(r['description']) for r in rows]))
