# encode=utf-8
import sys
import os
import re
import xlsxwriter


def main():
    txts = []
    for root, dirs, files in os.walk("."):
        for file in files:
            if (file.endswith(".txt")):
                txts.append(os.path.join(root, file))

    if len(txts) <= 0:
        print("没有找到任何txt文件")
        return

    coding = sys.argv[1] if len(sys.argv) > 1 else "gb2312"
    answer_file = "answer.dat"
    if not os.path.exists(answer_file):
        print("答案文件('%s')不存在" % answer_file)
        return

    def _readTxt(txt):
        with open(txt, "r", encoding=coding) as file:
            lines = file.readlines()[1:]
            answer_single, answer_multiple = [], []
            answer = answer_single
            for line in lines:
                if line.strip() == "多选":
                    answer = answer_multiple
                    continue

                mh = re.match(r"(\d+)([a-zA-Z]+)", line.strip())
                if not mh:
                    print("格式错误 %s : '%s'" % (txt, line.rstrip("\r\n")))
                    continue

                answer.append(
                    {"no": mh.group(1), "answer": mh.group(2).upper()})
            return answer_single, answer_multiple

    answer_single, answer_multiple = _readTxt(answer_file)
    grades = []
    for txt in txts:
        stu_grade = 0
        stu_single, stu_multiple = _readTxt(txt)
        for question in stu_single:
            answer = list(
                filter(lambda que: que["no"] == question["no"], answer_single))
            if answer and answer.pop()["answer"].upper() == question["answer"]:
                stu_grade += 3
        for question in stu_multiple:
            answer = list(
                filter(lambda que: que["no"] == question["no"], answer_multiple))
            if answer and answer.pop()["answer"].upper() == question["answer"]:
                stu_grade += 4

        grades.append({"name": os.path.splitext(
            os.path.basename(txt))[0], "grade": stu_grade})

    with xlsxwriter.Workbook("report.xlsx") as book:
        sheet = book.add_worksheet()
        sheet.write("A1", "姓名")
        sheet.write("B1", "成绩")
        row_num = 2
        for g in grades:
            sheet.write("A{}".format(row_num), g["name"])
            sheet.write("B{}".format(row_num), g["grade"])
            row_num += 1

    print("Done")


if __name__ == '__main__':
    main()
