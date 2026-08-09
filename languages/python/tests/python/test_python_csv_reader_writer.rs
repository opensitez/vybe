use super::helpers::run_python;

#[test]
fn test_python_csv_reader_rows() {
    let src = r#"
import csv
from io import StringIO

src = StringIO('a,b\n1,2\n3,4\n')
rows = list(csv.reader(src))
print(rows)
"#;
    assert_eq!(
        run_python(src),
        vec!["[['a', 'b'], ['1', '2'], ['3', '4']]"]
    );
}

#[test]
fn test_python_csv_writer_output() {
    let src = r#"
import csv
from io import StringIO

out = StringIO()
w = csv.writer(out)
w.writerow(['x', 'y'])
w.writerow([1, 2])
print(out.getvalue())
"#;
    assert_eq!(run_python(src), vec!["x,y\r\n1,2\r\n"]);
}
