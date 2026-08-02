# vybe-test: python/py_data_formats/test_py_csv_dictreader_dictwriter
# origin: languages/python/tests/python/test_py_data_formats.rs

import csv, io

buf = io.StringIO()
fieldnames = ["product", "price", "qty"]
writer = csv.DictWriter(buf, fieldnames=fieldnames)
writer.writeheader()
writer.writerow({"product": "apple", "price": 1.5, "qty": 100})
writer.writerow({"product": "banana", "price": 0.8, "qty": 200})

buf.seek(0)
reader = csv.DictReader(buf)
for row in reader:
    print(dict(row))
