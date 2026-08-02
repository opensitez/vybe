# vybe-test: python/python_csv_dialects_sniffer/test_csv_dict_reader_restkey_extra_fields
# origin: languages/python/tests/python/test_python_csv_dialects_sniffer.rs

import csv, io
data = "a,b\n1,2,3,4\n"
reader = csv.DictReader(io.StringIO(data), restkey="extra")
row = next(reader)
print(row["a"], row["b"], row["extra"])
