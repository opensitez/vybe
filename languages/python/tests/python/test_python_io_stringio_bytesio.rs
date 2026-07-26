// Python io.StringIO and io.BytesIO — in-memory I/O streams
use super::helpers::run_python;

#[test]
fn test_stringio_write_read() {
    let script = r#"
import io
buf = io.StringIO()
buf.write("hello ")
buf.write("world")
buf.seek(0)
print(buf.read())
"#;
    assert_eq!(run_python(script), vec!["hello world"]);
}

#[test]
fn test_stringio_getvalue() {
    let script = r#"
import io
buf = io.StringIO()
buf.write("abc\ndef\n")
print(buf.getvalue())
"#;
    assert_eq!(run_python(script), vec!["abc", "def", ""]);
}

#[test]
fn test_stringio_readline() {
    let script = r#"
import io
buf = io.StringIO("first\nsecond\nthird\n")
print(buf.readline().strip())
print(buf.readline().strip())
"#;
    assert_eq!(run_python(script), vec!["first", "second"]);
}

#[test]
fn test_bytesio_write_read() {
    let script = r#"
import io
buf = io.BytesIO()
buf.write(b'\x01\x02\x03')
buf.seek(0)
data = buf.read()
print(list(data))
"#;
    assert_eq!(run_python(script), vec!["[1, 2, 3]"]);
}

#[test]
fn test_bytesio_getvalue() {
    let script = r#"
import io
buf = io.BytesIO()
buf.write(b"ABC")
print(buf.getvalue())
"#;
    assert_eq!(run_python(script), vec!["b'ABC'"]);
}

#[test]
fn test_stringio_initial_value() {
    let script = r#"
import io
buf = io.StringIO("initial content")
print(buf.read())
"#;
    assert_eq!(run_python(script), vec!["initial content"]);
}

#[test]
fn test_stringio_seek_tell() {
    let script = r#"
import io
buf = io.StringIO("hello")
buf.read(3)
print(buf.tell())
buf.seek(0)
print(buf.tell())
"#;
    assert_eq!(run_python(script), vec!["3", "0"]);
}

#[test]
fn test_stringio_as_context_manager() {
    let script = r#"
import io
with io.StringIO() as buf:
    buf.write("test")
    buf.seek(0)
    print(buf.read())
"#;
    assert_eq!(run_python(script), vec!["test"]);
}
