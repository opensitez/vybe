# vybe-test: python/py_data_formats/test_py_csv_write_and_read
# origin: languages/python/tests/python/test_py_data_formats.rs

import csv, io

buf = io.StringIO()
writer = csv.writer(buf)
writer.writerow(["name", "age", "city"])
writer.writerow(["Alice", 30, "London"])
writer.writerow(["Bob", 25, "Paris"])

buf.seek(0)
reader = csv.reader(buf)
for row in reader:
    print(row)
