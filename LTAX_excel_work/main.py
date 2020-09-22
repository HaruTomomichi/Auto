import os
import openpyxl

class AUTO_SYSTEM():
    def __init__(self):
        self.a_sum = [0, 0]
        self.b_sum = [0, 0]
        self.c_sum = [0, 0]
        self.d_sum = [0, 0]
        self.e_sum = [0, 0]
        self.f_sum = [0, 0]
        self.g_sum = [0, 0]




if __name__=="__main__" :

    while 1:
        f_name = input("file name : ")

        cur_dir = os.getcwd()
        file_name = cur_dir + '/' + f_name
        wb = openpyxl.load_workbook(file_name)
        sheet = wb['(진단실행)진단항목오류정보']

        result = openpyxl.load_workbook(cur_dir + '/LTAX_값진단종합결과_테이블별오류취합_템플릿_200915_v0.01.xlsx')
        output = result['Sheet1']

