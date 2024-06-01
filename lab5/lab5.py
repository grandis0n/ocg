import argparse
import xlsxwriter
import csv

GROUP_NAME_INDEX = 0

parser = argparse.ArgumentParser(
    description="convert csv to xls"
)

parser.add_argument('input', type=str)
parser.add_argument('output', type=str)
parser.add_argument('-d', '--delimiter', default=';', type=str)

args = parser.parse_args()

csv_file = open(args.input, encoding='utf8', mode='r')
reader = csv.reader(csv_file)

workbook = xlsxwriter.Workbook(args.output)
worksheets = {}
records_in_group = {}
marks_count_in_group = {}
task_marks_sum_in_group = {}

header = next(reader)

for data in reader:
    group = data[GROUP_NAME_INDEX]

    if group not in worksheets:
        new_worksheet = workbook.add_worksheet(group)
        worksheets[group] = new_worksheet
        records_in_group[group] = 0
        marks_count_in_group[group] = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0}
        task_marks_sum_in_group[group] = [0 for _ in range(7)]

    row = records_in_group[group]

    points_sum = 0

    for col in range(1, len(data)):
        if data[col].isdigit():
            val = int(data[col])
        elif ',' in data[col]:
            val = float(data[col].replace(',', '.'))
        else:
            val = data[col]

        if isinstance(val, int) or isinstance(val, float):
            points_sum += val
            task_marks_sum_in_group[group][col - 2] += val

        worksheets[group].write(row, col - 1, val)

    rating_col = len(data) - 1
    rating = points_sum / 7.0
    worksheets[group].write(row, col, rating)

    mark_col = rating_col + 1

    if rating < 40:
        mark = 1
    elif rating < 50:
        mark = 2
    elif rating < 70:
        mark = 3
    elif rating < 80:
        mark = 4
    else:
        mark = 5

    marks_count_in_group[group][mark] += 1

    worksheets[group].write(row, mark_col, mark)

    records_in_group[group] += 1

total_records_count = reader.line_num - 1

for group in worksheets.keys():
    worksheet = worksheets[group]
    total_records_in_group = records_in_group[group]

    for i in range(7):
        avg_mark = task_marks_sum_in_group[group][i] / total_records_in_group
        worksheet.write(total_records_in_group + i + 1, 10, f'task{i + 1}')
        worksheet.write(total_records_in_group + i + 1, 11, avg_mark)

    chart = workbook.add_chart({'type': 'column'})
    start_row, start_col, end_row, end_col = total_records_in_group + 1, 11, total_records_in_group + 7, 11
    formula = f'={worksheet.name}!${xlsxwriter.utility.xl_rowcol_to_cell(start_row, start_col)}:${xlsxwriter.utility.xl_rowcol_to_cell(end_row, end_col)}'
    chart.add_series({'values': formula,
                      'categories': f'={worksheet.name}!${xlsxwriter.utility.xl_rowcol_to_cell(start_row, start_col - 1)}:${xlsxwriter.utility.xl_rowcol_to_cell(end_row, start_col - 1)}'})
    worksheet.insert_chart(total_records_in_group + 8, 10, chart)

    worksheet.autofit()

    marks_count = marks_count_in_group[group]

    for key in marks_count.keys():
        worksheet.write(total_records_in_group + key, 0, key)
        worksheet.write(total_records_in_group + key, 1, marks_count[key])

    pie_chart = workbook.add_chart({'type': 'pie'})

    start_row, start_col, end_row, end_col = total_records_in_group + 1, 1, total_records_in_group + 5, 1

    pie_formula = f'={worksheet.name}!${xlsxwriter.utility.xl_rowcol_to_cell(start_row, start_col)}:${xlsxwriter.utility.xl_rowcol_to_cell(end_row, end_col)}'
    pie_chart.add_series({'values': pie_formula})

    worksheet.insert_chart(total_records_in_group + 6, 0, pie_chart)

workbook.close()
csv_file.close()
