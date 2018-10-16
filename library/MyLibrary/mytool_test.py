import unittest

from mytool import mytool

class TestMytool(unittest.TestCase):

    # def test_init(self):
    #     my = mytool(self)

    def test_prepare_sql(self):
        my = mytool()
        filename = r'C:\Python27\Lib\site-packages\MyLibrary\test2.xls'
        sql = my.prepare_sql(filename)
        print sql

    def test_put_date_to_cell(self):
        my=mytool()
        filename = r'C:\Python27\Lib\site-packages\MyLibrary\test2.xls'
        date='2018/8/13'
        my.put_date_to_excel(filename,'tb_user',5,4,date)

    def test_get_value_from_excel(self):
        my=mytool()
        filename = r'C:\Python27\Lib\site-packages\MyLibrary\test2.xls'
        value = my.get_value_from_excel(filename,'tb_user',1,4)
        print value

    def test_put_string_to_cell(self):
        my=mytool()
        filename = r'C:\Python27\Lib\site-packages\MyLibrary\test2.xls'
        string='abc'
        my.put_string_to_excel(filename,'tb_user',1,2,string)

    if __name__ =='__main__':
        unittest.main()

