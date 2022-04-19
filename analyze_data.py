from collections import defaultdict
from xlutils.copy import copy
import xlrd, time


class AnalyzeData:
    FILE_PATH = 'C:/Users/MSI/Desktop/ex/data_copy.xls'  # 文件路径
    SCORE_TYPE_MAP = {  # 题目组
        1: {'col_num': range(14, 41), 'score_role': 1, 'desc': '生活质量'},
        2: {'col_num': range(41, 48), 'score_role': 1, 'desc': '焦虑'},
        3: {'col_num': range(48, 55), 'score_role': 1, 'desc': '抑郁'},
    }
    SCORE_ROLE = {  # 得分规则
        1: {1: 0, 2: 1, 3: 2, 4: 3}
    }

    def __init__(self):
        self.open_file()
        self.load_score_role = self.load_score_role()
        self.analyze_data = self.read_file()

    def open_file(self):
        """读取excel表格获取到对象"""
        self.excel_obj = xlrd.open_workbook(self.FILE_PATH, encoding_override="utf-8")
        self.new_path = self.get_new_file_path()
        self.sheet_obj = self.excel_obj.sheet_by_index(0)

    def get_new_file_path(self):
        path = self.FILE_PATH.split('/')
        path_title = '/'.join(path[0:-1])

        name = path[-1]
        name, file_format = name.split('.')

        return f"{path_title}/{name}_res_{int(time.time())}.{file_format}"

    def read_file(self):
        """读取文件所有内容"""
        data = defaultdict(list)
        analyze_data = {}
        for row_num in range(0, self.sheet_obj.nrows):
            if row_num == 0:
                max_col_num = self.sheet_obj.ncols
                for key, value in self.SCORE_TYPE_MAP.items():
                    key = (row_num, max_col_num + key)
                    analyze_data[key] = value['desc']

                continue
            data[row_num] = []
            row_group_score = {key: 0 for key in self.SCORE_TYPE_MAP.keys()}
            for col_num in range(0, self.sheet_obj.ncols):
                cell_data = self.sheet_obj.cell(row_num, col_num).value
                if col_num in self.load_score_role.keys():
                    print(row_num, col_num)
                    if int(cell_data) in self.load_score_role[col_num]['score_role']:
                        row_group_score[self.load_score_role[col_num]['group']] += self.load_score_role[col_num]['score_role'][int(cell_data)]
                data[row_num].append(cell_data)
            else:
                max_col_num = self.sheet_obj.ncols
                for key, score in row_group_score.items():
                    a_key = (row_num, max_col_num + key)
                    analyze_data[a_key] = score
        return analyze_data

    def load_score_role(self):
        col_score_role = {}
        for key, value in self.SCORE_TYPE_MAP.items():
            score_role = self.SCORE_ROLE[value['score_role']]
            for col_num in value['col_num']:
                col_score_role[col_num] = {
                    'score_role': score_role,
                    'group': key
                }
        print(col_score_role)
        return col_score_role

    def write_data(self):
        excel_file_copy = copy(self.excel_obj)
        new_sheet = excel_file_copy.get_sheet(0)
        for key, value in self.analyze_data.items():
            new_sheet.write(key[0], key[1], value)
        excel_file_copy.save(self.new_path)


if __name__ == '__main__':
    analyze_data = AnalyzeData()
    analyze_data.write_data()
