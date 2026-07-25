use super::helpers::run_python;

// csv — DictReader, DictWriter, Sniffer, QUOTE_ALL, QUOTE_MINIMAL, QUOTE_NONNUMERIC, QUOTE_NONE, register_dialect, fieldnames, restval, restkey

#[test]
fn test_csv_dict_reader_basic() {
    let out = run_python(r#"
import csv, io
data = "name,age,city\nAlice,30,NY\nBob,25,LA\n"
reader = csv.DictReader(io.StringIO(data))
rows = list(reader)
print(rows[0]["name"], rows[0]["age"])
print(rows[1]["name"], rows[1]["city"])
"#);
    assert_eq!(out, vec!["Alice 30", "Bob LA"]);
}

#[test]
fn test_csv_dict_writer_basic() {
    let out = run_python(r#"
import csv, io
buf = io.StringIO()
writer = csv.DictWriter(buf, fieldnames=["id", "name", "role"])
writer.writeheader()
writer.writerow({"id": "1", "name": "Admin", "role": "superuser"})
print(buf.getvalue().strip())
"#);
    assert_eq!(out, vec!["id,name,role\n1,Admin,superuser"]);
}

#[test]
fn test_csv_quote_nonnumeric_writer() {
    let out = run_python(r#"
import csv, io
buf = io.StringIO()
writer = csv.writer(buf, quoting=csv.QUOTE_NONNUMERIC)
writer.writerow(["text", 42, 3.14, True])
print(buf.getvalue().strip())
"#);
    assert_eq!(out, vec!["\"text\",42,3.14,1"]);
}

#[test]
fn test_csv_quote_all_writer() {
    let out = run_python(r#"
import csv, io
buf = io.StringIO()
writer = csv.writer(buf, quoting=csv.QUOTE_ALL)
writer.writerow([10, "hello", 20.5])
print(buf.getvalue().strip())
"#);
    assert_eq!(out, vec!["\"10\",\"hello\",\"20.5\""]);
}

#[test]
fn test_csv_quote_none_with_escapechar() {
    let out = run_python(r#"
import csv, io
buf = io.StringIO()
writer = csv.writer(buf, quoting=csv.QUOTE_NONE, escapechar="\\")
writer.writerow(["hello, world", "foo:bar"])
print(buf.getvalue().strip())
"#);
    assert_eq!(out, vec!["hello\\, world,foo:bar"]);
}

#[test]
fn test_csv_register_custom_dialect() {
    let out = run_python(r#"
import csv, io
csv.register_dialect("pipes", delimiter="|", quoting=csv.QUOTE_MINIMAL)
buf = io.StringIO()
writer = csv.writer(buf, dialect="pipes")
writer.writerow(["a", "b", "c"])
print(buf.getvalue().strip())
csv.unregister_dialect("pipes")
"#);
    assert_eq!(out, vec!["a|b|c"]);
}

#[test]
fn test_csv_dict_writer_restval_default() {
    let out = run_python(r#"
import csv, io
buf = io.StringIO()
writer = csv.DictWriter(buf, fieldnames=["a", "b", "c"], restval="N/A")
writer.writerow({"a": "val1"})
print(buf.getvalue().strip())
"#);
    assert_eq!(out, vec!["val1,N/A,N/A"]);
}

#[test]
fn test_csv_dict_reader_restkey_extra_fields() {
    let out = run_python(r#"
import csv, io
data = "a,b\n1,2,3,4\n"
reader = csv.DictReader(io.StringIO(data), restkey="extra")
row = next(reader)
print(row["a"], row["b"], row["extra"])
"#);
    assert_eq!(out, vec!["1", "2", "['3', '4']"]);
}

#[test]
fn test_csv_sniffer_sniff_delimiter() {
    let out = run_python(r#"
import csv
sample = "name;age;city\nAlice;30;London\nBob;22;Paris\n"
dialect = csv.Sniffer().sniff(sample)
print(dialect.delimiter)
"#);
    assert_eq!(out, vec![";"]);
}

#[test]
fn test_csv_sniffer_has_header() {
    let out = run_python(r#"
import csv
header_sample = "header1,header2,header3\n1,2,3\n4,5,6\n"
no_header_sample = "1,2,3\n4,5,6\n7,8,9\n"
print(csv.Sniffer().has_header(header_sample))
print(csv.Sniffer().has_header(no_header_sample))
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_csv_dict_writer_extrasaction_ignore() {
    let out = run_python(r#"
import csv, io
buf = io.StringIO()
writer = csv.DictWriter(buf, fieldnames=["x", "y"], extrasaction="ignore")
writer.writerow({"x": 1, "y": 2, "ignored_key": 99})
print(buf.getvalue().strip())
"#);
    assert_eq!(out, vec!["1,2"]);
}

#[test]
fn test_csv_dict_writer_extrasaction_raise() {
    let out = run_python(r#"
import csv, io
buf = io.StringIO()
writer = csv.DictWriter(buf, fieldnames=["x"], extrasaction="raise")
try:
    writer.writerow({"x": 1, "extra": 2})
except ValueError:
    print("ValueError")
"#);
    assert_eq!(out, vec!["ValueError"]);
}

#[test]
fn test_csv_field_size_limit() {
    let out = run_python(r#"
import csv
old_limit = csv.field_size_limit(1000000)
print(isinstance(old_limit, int))
print(csv.field_size_limit())
csv.field_size_limit(old_limit)  # restore
"#);
    assert_eq!(out, vec!["True", "1000000"]);
}

#[test]
fn test_csv_dict_reader_fieldnames_override() {
    let out = run_python(r#"
import csv, io
data = "val1,val2\nval3,val4\n"
reader = csv.DictReader(io.StringIO(data), fieldnames=["col1", "col2"])
rows = list(reader)
print(rows[0]["col1"], rows[0]["col2"])
print(rows[1]["col1"], rows[1]["col2"])
"#);
    assert_eq!(out, vec!["val1 val2", "val3 val4"]);
}

#[test]
fn test_csv_writer_writerows() {
    let out = run_python(r#"
import csv, io
buf = io.StringIO()
writer = csv.writer(buf)
writer.writerows([["r1c1", "r1c2"], ["r2c1", "r2c2"]])
print(buf.getvalue().strip().replace("\r\n", "\n"))
"#);
    assert_eq!(out, vec!["r1c1,r1c2\nr2c1,r2c2"]);
}

#[test]
fn test_csv_reader_line_num() {
    let out = run_python(r#"
import csv, io
data = "line1\nline2\nline3\n"
reader = csv.reader(io.StringIO(data))
nums = []
for row in reader:
    nums.append(reader.line_num)
print(nums)
"#);
    assert_eq!(out, vec!["[1, 2, 3]"]);
}

#[test]
fn test_csv_dict_writer_writeheader() {
    let out = run_python(r#"
import csv, io
buf = io.StringIO()
writer = csv.DictWriter(buf, fieldnames=["alpha", "beta"])
header_str = writer.writeheader()
print(buf.getvalue().strip())
"#);
    assert_eq!(out, vec!["alpha,beta"]);
}

#[test]
fn test_csv_excel_dialect_default() {
    let out = run_python(r#"
import csv
dialect = csv.get_dialect("excel")
print(dialect.delimiter)
print(dialect.doublequote)
print(dialect.lineterminator.encode())
"#);
    assert_eq!(out, vec![",", "True", "b'\\r\\n'"]);
}

#[test]
fn test_csv_excel_tab_dialect() {
    let out = run_python(r#"
import csv
dialect = csv.get_dialect("excel-tab")
print(dialect.delimiter)
"#);
    assert_eq!(out, vec!["\t"]);
}

#[test]
fn test_csv_list_dialects() {
    let out = run_python(r#"
import csv
dialects = csv.list_dialects()
print("excel" in dialects)
print("excel-tab" in dialects)
"#);
    assert_eq!(out, vec!["True", "True"]);
}
