# vybe-test: python/python_csv_dialects_sniffer/test_csv_reader_line_num
# origin: languages/python/tests/python/test_python_csv_dialects_sniffer.rs

import csv, io
data = "line1\nline2\nline3\n"
reader = csv.reader(io.StringIO(data))
nums = []
for row in reader:
    nums.append(reader.line_num)
print(nums)
