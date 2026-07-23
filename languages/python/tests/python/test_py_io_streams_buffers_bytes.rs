use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: IO Streams & Buffer Operations — io.StringIO, io.BytesIO, TextIOWrapper, seek, tell, read, write
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_io_stringio_read_write_seek() {
    let src = r#"
import io

buf = io.StringIO()
buf.write("Hello ")
buf.write("World")
buf.seek(0)
print(buf.read())
print(buf.getvalue())
"#;
    assert_eq!(run_python(src), vec!["Hello World", "Hello World"]);
}

#[test]
fn test_py_io_bytesio_binary_buffer() {
    let src = r#"
import io

buf = io.BytesIO(b"\x00\x01\x02\x03")
print(buf.read(2))
print(buf.tell())
buf.seek(0, io.SEEK_END)
print(buf.tell())
"#;
    assert_eq!(run_python(src), vec!["b'\\x00\\x01'", "2", "4"]);
}

#[test]
fn test_py_io_textiowrapper_encoding_wrap() {
    let src = r#"
import io

raw_buf = io.BytesIO(b"Hello UTF-8 \xc3\xa9")
text_stream = io.TextIOWrapper(raw_buf, encoding="utf-8")
print(text_stream.read())
"#;
    assert_eq!(run_python(src), vec!["Hello UTF-8 é"]);
}

#[test]
fn test_py_io_stringio_truncate_buffer() {
    let src = r#"
import io

buf = io.StringIO("long initial content")
buf.seek(4)
buf.truncate()
print(repr(buf.getvalue()))
"#;
    assert_eq!(run_python(src), vec!["'long'"]);
}

#[test]
fn test_py_io_line_iteration_readlines() {
    let src = r#"
import io

data = "line1\nline2\nline3"
buf = io.StringIO(data)
lines = [line.strip() for line in buf]
print(lines)
"#;
    assert_eq!(run_python(src), vec!["['line1', 'line2', 'line3']"]);
}

#[test]
fn test_py_io_bytesio_getbuffer_zero_copy() {
    let src = r#"
import io

buf = io.BytesIO(b"ABCDEF")
view = buf.getbuffer()
print(view[0])
print(bytes(view[1:4]).decode())
"#;
    assert_eq!(run_python(src), vec!["65", "BCD"]);
}

#[test]
fn test_py_io_buffered_reader_writer() {
    let src = r#"
import io

raw = io.BytesIO()
writer = io.BufferedWriter(raw)
writer.write(b"buffered data")
writer.flush()
print(raw.getvalue().decode())
"#;
    assert_eq!(run_python(src), vec!["buffered data"]);
}

#[test]
fn test_py_io_seek_relative_constants() {
    let src = r#"
import io

buf = io.BytesIO(b"0123456789")
buf.seek(5, io.SEEK_SET)
print(buf.read(2).decode())

buf.seek(-3, io.SEEK_END)
print(buf.read().decode())
"#;
    assert_eq!(run_python(src), vec!["56", "789"]);
}

#[test]
fn test_py_io_closed_file_operations_error() {
    let src = r#"
import io

buf = io.StringIO("test")
buf.close()
print(buf.closed)
try:
    buf.read()
except ValueError as e:
    print("ValueError: I/O operation on closed file")
"#;
    assert_eq!(
        run_python(src),
        vec!["True", "ValueError: I/O operation on closed file"]
    );
}

#[test]
fn test_py_io_incremental_newline_translation() {
    let src = r#"
import io

raw = io.BytesIO(b"line1\r\nline2\rline3\n")
text = io.TextIOWrapper(raw, newline=None)
lines = text.read().splitlines()
print(lines)
"#;
    assert_eq!(run_python(src), vec!["['line1', 'line2', 'line3']"]);
}
