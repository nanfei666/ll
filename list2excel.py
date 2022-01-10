import re
import xlwt
import xlsxwriter

class My_xlwt(object):
    def __init__(self,name,sheet_name = 'sheet_1',re_write = True):
        '''
        自定义类说明：
        :param sheet_name:默认sheet表对象名称，默认值为 'sheet_1'
        :param re_write: 单元格重写写功能默认开启
        '''
        self.work_book = xlsxwriter.Workbook(name)
        self.sheet = self.work_book.add_worksheet(sheet_name)
        self.col_data = {}
    def save(self):
        self.work_book.close()

    def write(self,row,col,label):
        '''
        在默认sheet表对象一个单元格内写入数据
        :param row: 写入行
        :param col: 写入列
        :param label: 写入数据
        '''
        self.sheet.write(row,col,label)
        
        # 将列数据加入到col_data字典中
        if col not in self.col_data.keys():
            self.col_data[col] = []
            self.col_data[col].append(label)
        else:
            self.col_data[col].append(label)

    def write_row(self,start_row,start_col,date_list):
        '''
        按行写入一行数据
        :param start_row:写入行序号
        :param start_col: 写入列序号
        :param date_list: 写入数据：列表
        :return: 返回行对象
        '''
        for col,label in enumerate(date_list):
            if label==[]:
                self.write(start_row,start_col+col,'')
            else:
                self.write(start_row,start_col+col,label)

def main(txt_file,excel_name):
    with open (txt_file,'r') as f:
        result=[]
        for i,line in enumerate(f.readlines()):
            reg=r'\s-?[1-9]+[0-9]*.?[0-9]*E-?\+?[0-9]+\s?'

            target = re.findall(reg,line.strip())
            floatAr  = [x for x in target]
            if floatAr:
                result.append(floatAr)

        test = My_xlwt(name=excel_name)
        j=0
        for i in range(len(result)):
            # 在0行0列写入一行数据
            j+=1
            if j%14==0:
                test.write_row(i,0,[])
                j=0
            else:
                test.write_row(i,0,result[i])
            
            # 保存文件
        test.save()

txt_file='only_push_No9_lonGMI.txt'
excel_name = 'my_test.xls'
main(txt_file,excel_name)
