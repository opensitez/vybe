use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: io — StringIO, BytesIO, file reading, writing, seeking, tempfile, text vs binary modes
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_io_stringio_write_and_read() {
    let src = r#"
import io

buf = io.StringIO()
buf.write("Hello, ")
buf.write("world!")
buf.seek(0)
print(buf.read())
print(buf.getvalue())
"#;
    assert_eq!(run_python(src), vec!["Hello, world!", "Hello, world!"]);
}

#[test]
fn test_py_io_stringio_readline() {
    let src = r#"
import io

buf = io.StringIO("line1\nline2\nline3\n")
print(buf.readline().strip())
print(buf.readline().strip())
print(buf.readlines())
"#;
    assert_eq!(run_python(src), vec!["line1", "line2", "['line3\\n']"]);
}

#[test]
fn test_py_io_bytesio_write_and_read() {
    let src = r#"
import io

buf = io.BytesIO()
buf.write(b"binary data")
buf.seek(0)
print(buf.read())
print(buf.tell())
"#;
    assert_eq!(run_python(src), vec!["b'binary data'", "11"]);
}

#[test]
fn test_py_io_bytesio_seek_and_tell() {
    let src = r#"
import io

buf = io.BytesIO(b"ABCDEFGHIJ")
buf.seek(3)
print(buf.read(4))
print(buf.tell())
buf.seek(-2, 2)  # 2 bytes from end
print(buf.read())
"#;
    assert_eq!(run_python(src), vec!["b'DEFG'", "7", "b'IJ'"]);
}

#[test]
fn test_py_io_write_and_read_text_file() {
    let src = r#"
import tempfile, os

with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
    fname = f.name
    f.write("first line\nsecond line\nthird line\n")

with open(fname, 'r') as f:
    lines = f.readlines()
    print(len(lines))
    print(lines[0].strip())

os.unlink(fname)
"#;
    assert_eq!(run_python(src), vec!["3", "first line"]);
}

#[test]
fn test_py_io_read_binary_file() {
    let src = r#"
import tempfile, os

with tempfile.NamedTemporaryFile(delete=False) as f:
    fname = f.name
    f.write(b'\x00\x01\x02\x03')

with open(fname, 'rb') as f:
    data = f.read()
    print(len(data))
    print(data[2])

os.unlink(fname)
"#;
    assert_eq!(run_python(src), vec!["4", "2"]);
}

#[test]
fn test_py_io_append_mode() {
    let src = r#"
import tempfile, os

with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
    fname = f.name
    f.write("initial\n")

with open(fname, 'a') as f:
    f.write("appended\n")

with open(fname) as f:
    lines = f.readlines()
    print([l.strip() for l in lines])

os.unlink(fname)
"#;
    assert_eq!(run_python(src), vec!["['initial', 'appended']"]);
}

#[test]
fn test_py_io_seek_tell_truncate() {
    let src = r#"
import tempfile, os

with tempfile.NamedTemporaryFile(mode='w+', suffix='.txt', delete=False) as f:
    fname = f.name
    f.write("Hello World")
    f.seek(0)
    f.truncate(5)
    f.seek(0)
    print(f.read())

os.unlink(fname)
"#;
    assert_eq!(run_python(src), vec!["Hello"]);
}

#[test]
fn test_py_io_tempfile_namedtemporary() {
    let src = r#"
import tempfile

with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=True) as f:
    f.write('{"test": true}')
    f.flush()
    print(f.name.endswith('.json'))
    print(f.mode)
"#;
    assert_eq!(run_python(src), vec!["True", "w"]);
}

#[test]
fn test_py_io_temporarydirectory() {
    let src = r#"
import tempfile, os
from pathlib import Path

with tempfile.TemporaryDirectory() as tmpdir:
    p = Path(tmpdir) / "test.txt"
    p.write_text("contents")
    print(p.read_text())
    print(os.path.isdir(tmpdir))

print(os.path.isdir(tmpdir))  # cleaned up after exit
"#;
    assert_eq!(run_python(src), vec!["contents", "True", "False"]);
}

#[test]
fn test_py_io_buffered_reader_writer() {
    let src = r#"
import io

raw = io.BytesIO(b"Hello, World!")
buffered = io.BufferedReader(raw)
data = buffered.read(5)
print(data)
print(buffered.read())
"#;
    assert_eq!(run_python(src), vec!["b'Hello'", "b', World!'"]);
}

#[test]
fn test_py_io_stringio_as_context_manager() {
    let src = r#"
import io

with io.StringIO() as buf:
    buf.write("Context managed!")
    buf.seek(0)
    print(buf.read())
"#;
    assert_eq!(run_python(src), vec!["Context managed!"]);
}

#[test]
fn test_py_io_encoding_errors_handling() {
    let src = r#"
import io

# Write UTF-8 bytes and read with different error handling
data = "café".encode("utf-8")
buf = io.BytesIO(data)

# Using TextIOWrapper with encoding
wrapped = io.TextIOWrapper(buf, encoding="utf-8")
print(wrapped.read())
"#;
    assert_eq!(run_python(src), vec!["café"]);
}
