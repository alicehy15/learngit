# coding=utf-8
import xlrd
import json
from datetime import datetime
from xlutils.copy import copy
import xlwt
from os.path import join
import sys
import struct
import requests
import codecs
from warnings import catch_warnings
reload(sys)
sys.setdefaultencoding('utf-8')

class mytool():
    def __init__(self):
        pass
    def test_a_b(self,a,b):
        '''
        比较两个参数的大小
        '''
        if a>b:
             flag = False
             return flag
        else:
             flag = True
             return flag

    def get_menures(self,filename,modelname,menuname):
        '''
        根据模块名和菜单名给出resource_id
        '''
        data = xlrd.open_workbook(filename)
        table = data.sheet_by_index(0)
        n = table.nrows
        for i in range(n):
            if table.cell_value(i,0) != modelname:
                continue;
            else:

                if table.cell_value(i,2) == menuname:
                    val=table.cell_value(i,1)
                    if(val==None or ''==val):
                        raise RuntimeError
                    else:
                        return val

        else:
            raise RuntimeError

    def get_data_by_file(self,filename):
        '''根据输入的Excel表格，第一行为key值，接下来为value值，返回字典的list'''
        data = xlrd.open_workbook(filename)
        table = data.sheet_by_index(0)
        n = table.nrows
        m = table.ncols
        dictlist = []
        eh = excelHandle()
        for j in range(n-1):
            # 产生一个数据dict
            keys = {}
            for i in range(m):
                keys[eh.read_cell_value(table,0,i)] = eh.read_cell_value(table,j+1,i)
            #将一次数据附加到数据列表中
            dictlist.append(keys)
        print dictlist
        print len(dictlist)
        return dictlist
    
    def export_to_excel(self,data,filename):
        '''输入response的content,并指定文件名，将二进制流传入文件中'''
        file = open(filename,"wb")
        file.write(data.content)
        file.close()
        return 0

    def get_value_from_excel(self,filename,sheetname,row,col):
        '''根据文件名，sheetname,行，列，返回数据'''
        data=xlrd.open_workbook(filename)
        my_sheet_index = data.sheet_names().index(sheetname)
        table = data.sheet_by_index(my_sheet_index)
        eh = excelHandle()
        value = eh.read_cell_value(table,int(row),int(col))
        return value

    def put_string_to_excel(self,filename,sheetname,row,col,value):
        '''根据文件名，sheetname,行，列，插入数据'''
        data = xlrd.open_workbook(filename=filename, formatting_info=True, on_demand=True)
        my_sheet_index = data.sheet_names().index(sheetname)
        w = copy(data)
        data.release_resources()
        style = xlwt.XFStyle()
        style.num_format_str = 'general'
        w.get_sheet(my_sheet_index).write(int(row),int(col),value,style)
        w.save(filename)
    
    def put_date_to_excel(self,filename,sheetname,row,col,value):
        '''根据文件名，sheetname,行，列，插入数据,datetime格式：yyyy/M/d例如 2018/8/13'''
        print filename
        # data = xlrd.open_workbook(filename=filename)
        data = xlrd.open_workbook(filename=filename,formatting_info=True,on_demand=True)
        my_sheet_index = data.sheet_names().index(sheetname)
        w = copy(data)
        # 释放资源
        data.release_resources()
        dt = value.split('/')
        dti = [int(dt[0]), int(dt[1]), int(dt[2])]
        print(dt, dti)
        ymd = datetime(*dti)
        style = xlwt.XFStyle()
        style.num_format_str = 'yyyy/M/d'
        w.get_sheet(my_sheet_index).write(int(row),int(col),ymd, style)
        # filename = r'C:\Python27\Lib\site-packages\MyLibrary\test.xls'
        w.save(filename)
    
    def prepare_sql(self,filename):
        """根据xls的内容，返回sql语句组成的列表
        
        Examples:
        | ${path} | Catenate | SEPARATOR=\\ | ${CURDIR} | test2.xls |
        | ${sqls} | Prepare Sql | ${path} |
        
        test2.xls
        tablename = sheetname
         第一行为参数名，剩下的为要插入的参数值
        """
        data = xlrd.open_workbook(filename)
        sheetnames = data.sheet_names()
        sql = []
        for sn in sheetnames:
            # 根据sheetname判断插入哪张表
            table = data.sheet_by_name(sn)
            # 根据表格内容拼接sql语句
            n = table.nrows
            m = table.ncols
            eh = excelHandle()
            keys = []
            for j in range(m):
                key='`'+eh.read_cell_value(table,0,j)+'`'
                keys.append(key)
            k = ','.join(str(n) for n in keys)
            for i in range(n-1):
                values = []
                for j in range(m):
                    value='\''+str(eh.read_cell_value(table, i+1, j))+'\''
                    values.append(value)
                v=','.join(str(n) for n in values)
                sql_insert = u'insert into `' + sn + u'` (' + k +u') values (' + v +u')'
                sql.append(sql_insert)
        return sql

    def get_json_from_file(self,path):
        '''从json文件中获取数据'''
        file = open(path,"rb")
        fileJson = json.load(file)
        return fileJson

    def diff(self,data1,data2):
        '''
        比较两个入参结果的一致性，并打印出不同
        '''
        if isinstance(data1, dict)&isinstance(data2, dict):
            result = cmp(data1,data2)
            if result != 0:
                keys = data1.keys()
                keys.extend(data2.keys())
                keys=set(keys)
                for i in keys:
                    if data1.get(i)!=data2.get(i):
                        return 'key=' + i + ',data1 value=' + str(data1.get(i)) + ',data2 value = ' + str(data2.get(i))
            else: return 0
        elif isinstance(data1, list)&isinstance(data2, list):
             result = cmp(data1.sort(),data2.sort())
             if result !=0:
                 return 1
        else:
            return u'不支持的比较格式'

    def set_value_from_list(self,data,keys):
        '''
        根据keys中的key值从data中提取出键值对组成一个新的dict
        '''
        content = {}
#         keys = data2.keys()
        for key in keys:
            try:
                value = data[key]
                content[key]=value
            except KeyError:
                print key
                continue
            except TypeError:
                print key
                continue
        return content

class excelHandle():
    def decode(self, filename, sheetname):
        try:
            filename = filename.decode('utf-8')
            sheetname = sheetname.decode('utf-8')
        except Exception:
            print traceback.print_exc()
        return filename, sheetname

    def read_excel(self, filename, sheetname):
        filename, sheetname = self.decode(filename, sheetname)
        rbook = xlrd.open_workbook(filename)
        sheet = rbook.sheet_by_name(sheetname)
        rows = sheet.nrows
        cols = sheet.ncols
        all_content = []
        for i in range(rows):
            row_content = []
            for j in range(cols):
                ctype = sheet.cell(i, j).ctype  # 表格的数据类型
                cell = sheet.cell_value(i, j)
                if ctype == 2 and cell % 1 == 0:  # 如果是整形
                    cell = int(cell)
                elif ctype == 3:
                    # 转成datetime对象
                    date = datetime(*xldate_as_tuple(cell, 0))
                    cell = date.strftime('%Y/%d/%m %H:%M:%S')
                elif ctype == 4:
                    cell = True if cell == 1 else False
                row_content.append(cell)
            all_content.append(row_content)
            print '[' + ','.join("'" + str(element) + "'" for element in row_content) + ']'
        return all_content

    def read_cell_value(self,table,row,col):
        ctype = table.cell(row,col).ctype
        cell = table.cell_value(row,col)
        if ctype == 1 or ctype == 0:
            cell = cell.encode('utf-8')
        elif ctype == 2 and cell % 1 == 0:  # 如果是整形
            cell = int(cell)
        elif ctype == 3:
            # 转成datetime对象
            date = datetime(*xldate_as_tuple(cell, 0))
            cell = date.strftime('%Y/%m/%d %H:%M:%S')
        elif ctype == 4:
            cell = True if cell == 1 else False
        return cell
