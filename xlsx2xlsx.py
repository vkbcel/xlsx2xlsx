from io import StringIO
from openpyxl import load_workbook, Workbook


def openfile(filename):
    wb = load_workbook(filename=filename, read_only=True)
    ws = wb.worksheets[0]
    temp = StringIO()

    for row in ws.rows:
        if row[0].value:
            temp.writelines(row[0].value + '\n')
    return temp


def handle(fd, out_filename):
    wb = Workbook(write_only=True)
    ws = wb.create_sheet()
    fd.seek(0)

    title = False

    line = fd.readline().split(',')
    for __ in range(1000):
        key, value, sched = [], [], []
        if line[0] == 'RELEASE HEADER SECTION':
            key, value = [], []
            key.extend(fd.readline().split(','))
            value.extend(fd.readline().split(','))

            line = fd.readline().split(',')
            key.extend(fd.readline().split(','))
            value.extend(fd.readline().split(','))

            line = fd.readline().split(',')
            key.extend(fd.readline().split(','))
            value.extend(fd.readline().split(','))

            line = fd.readline().split(',')
            key.extend(fd.readline().split(','))
            while True:
                line = fd.readline().split(',')
                if not line[0] or line[0] == 'RELEASE HEADER SECTION':
                    break
                sched.append(line)

        if not title:
            ws.append(key)
            title = True
        for i in range(len(sched)):
            row = []
            row.extend(value)
            row.extend(sched[i])
            if i > 0:
                row.append('DITTO')
            ws.append(row)

    wb.save(out_filename)


if __name__ == '__main__':
    a = openfile('download.xlsx')
    handle(a, 'out.xlsx')
