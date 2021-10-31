# -*- coding: utf-8 -*-
# @Time    : 2021/10/27 19:12
# @Author  : HanMengnan
# @Email   : hanmnan@163.com
# @FileName: test.py
# @Software: PyCharm

import xlrd
import xlwt


# 在进行时间运算时
# 将“02：10”这样由小时和分钟组成的时间转换为，以小数格式小时表示的时间
def parse_time(time_str):
    s = time_str.split(':')
    return round(float(s[0]) + float(s[1]) / 60, 2)


class Attendance:
    def __init__(self):
        self.file = xlwt.Workbook()
        self.sheet = self.file.add_sheet("sheet1", cell_overwrite_ok=True)
        # 目前进行写入的行数
        self.write_row = 0
        self.data = []

    def calculate(self, fileName, doctor=False):
        # 读取源数据，并改变编码格式
        xls = xlrd.open_workbook_xls(fileName, encoding_override="gbk")
        sheet = xls.sheet_by_index(0)
        name_item_list = sheet.col_slice(0, 1, sheet.nrows)

        # 每个名字对应的行数
        name_row_map = {}

        for index, item, in enumerate(name_item_list):
            if item.value not in name_row_map:
                name_row_map[item.value] = [sheet.row(index + 1)]
            else:
                name_row_map[item.value].append(sheet.row(index + 1))

        self.data = []

        # 遍历每个名字，计算该人本周的统计数据
        for key, value in name_row_map.items():
            p = PersonSummary(key, value, doctor)
            self.data.append(p.calculate())

        # 按照加权出勤时间进行排序
        self.sort_res()
        # 将本年纪数据写入excel
        self.write_res()
        # 两个年纪之间需要空一行
        self.write_row += 1

    def write_res(self):
        title = ["姓名", "应到", "实到", "迟到", "早退", "旷工", "加班", "工作时间", "未签到", "未签退", "出勤时间", "加权出勤"]
        late_style = xlwt.easyxf('pattern: pattern solid, fore_colour sky_blue')
        not_be_style = xlwt.easyxf('pattern: pattern solid, fore_colour light_orange')

        for index, item in enumerate(title):
            self.sheet.write(self.write_row, index, title[index])
        self.write_row += 1

        for stu in self.data:
            for index, item in enumerate(stu):
                if index == 3 and item != 0:
                    # 迟到，颜色标注
                    self.sheet.write(self.write_row, index, item, late_style)
                elif index == 5 and item != 0:
                    # 旷工，颜色标注
                    self.sheet.write(self.write_row, index, item, not_be_style)
                else:
                    self.sheet.write(self.write_row, index, item)
            self.write_row += 1

    def save_res(self):
        self.file.save("./打卡汇总.xls")

    def sort_res(self):
        def get_weight(elem):
            return elem[11]

        self.data.sort(reverse=True, key=get_weight)


class PersonSummary:
    def __init__(self, name, rows, doctor):
        self.name = name
        self.rows = rows

        self.should_be_num = 0
        self.be_num = 0
        self.late_be_num = 0
        self.early_leave_num = 0
        self.not_be_num = 0
        self.extra_hour = 0
        self.work_hour = 0
        self.not_record_be = 0
        self.not_record_leave = 0
        self.be_hour = 0
        self.weighting_be_hour = 0
        # 判断是否是博士
        if doctor:
            self.penalty_unit = 7
        else:
            self.penalty_unit = 6.5
        self.penalty_time = 0

    def calculate(self):
        for row in self.rows:
            # 应到
            self.should_be_num = self.should_be_num + float(row[4].value)
            # 实到
            if row[5].value != "":
                self.be_num = self.be_num + float(row[5].value)
            # 迟到, 迟到15分钟以内不计
            if row[6].value != "" and parse_time(row[6].value) > parse_time("00:15"):
                self.late_be_num = self.late_be_num + 0.5
                self.penalty_time = self.penalty_time + parse_time(row[6].value)
            # 早退
            if row[7].value != "":
                self.early_leave_num = self.early_leave_num + 0.5
                self.penalty_time = self.penalty_time + parse_time(row[7].value)
            # 旷工
            if row[2].value == "" and row[3].value == "":
                self.not_be_num = self.not_be_num + 0.5
            # 加班
            if row[9].value != "":
                self.extra_hour = self.extra_hour + parse_time(row[9].value)
            # 工作
            if row[10].value != "":
                self.work_hour = self.work_hour + parse_time(row[10].value)
            # 出勤
            if row[11].value != "":
                self.be_hour = self.be_hour + parse_time(row[11].value)
            # 未签到
            if row[2].value == "" and row[3].value != "":
                self.not_record_be = self.not_record_be + 0.5
                # 未签到影响计算工作时间和出勤时间
                if parse_time(row[3].value) < parse_time("16:00"):
                    self.work_hour = self.work_hour + (parse_time("17:30") - parse_time("14:00"))
                    self.be_hour = self.be_hour + (parse_time(row[3].value) - parse_time("14:00"))
                else:
                    if self.penalty_time == 7:
                        self.work_hour = self.work_hour + (parse_time("12:00") - parse_time("8:30"))
                        self.be_hour = self.be_hour + (parse_time(row[3].value) - parse_time("8:30"))
                    else:
                        self.work_hour = self.work_hour + (parse_time("12:00") - parse_time("9:00"))
                        self.be_hour = self.be_hour + (parse_time(row[3].value) - parse_time("9:00"))

            # 未签退
            if row[3].value == "" and row[2].value != "":
                self.not_record_leave += self.not_record_leave + 0.5
                # 未签到影响计算工作时间和出勤时间
                if parse_time(row[2].value) > parse_time("12:00"):
                    self.work_hour = self.work_hour + (parse_time("17:30") - parse_time("14:00"))
                    self.be_hour = self.be_hour + (parse_time("17:30") - parse_time(row[2].value))
                else:
                    self.work_hour = self.work_hour + (parse_time("12:00") - parse_time("8:30"))
                    self.be_hour = self.be_hour + (parse_time("12:00") - parse_time(row[2].value))

        # 加权出勤
        # 出勤时间 + 加班时间 - 旷工 * 每天打卡时间 - 早退 - 迟到
        self.weighting_be_hour = self.be_hour + self.extra_hour - self.penalty_time - (
                self.penalty_unit * self.not_be_num)

        return [self.name, self.should_be_num, self.be_num, self.late_be_num, self.early_leave_num, self.not_be_num,
                round(self.extra_hour, 2), round(self.work_hour, 2), self.not_record_be, self.not_record_leave,
                round(self.be_hour, 2),
                round(self.weighting_be_hour, 2)]


if __name__ == "__main__":
    att = Attendance()
    att.calculate("./1.xls")
    att.calculate("./2.xls")
    att.calculate("./3.xls")
    att.save_res()
