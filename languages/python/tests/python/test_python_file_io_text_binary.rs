// Python file I/O — text mode, binary mode, readline, writelines, seek, tell
use super::helpers::run_python;

#[test]
fn test_write_and_read_text_file() {
    let script = r#"
import tempfile, os

with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
    path = f.name
    f.write("hello\nworld\n")

with open(path, 'r') as f:
    content = f.read()

os.unlink(path)
print(content)
"#;
    assert_eq!(run_python(script), vec!["hello", "world", ""]);
}

#[test]
fn test_readline_and_readlines() {
    let script = r#"
import tempfile, os

with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
    path = f.name
    f.writelines(["line1\n", "line2\n", "line3\n"])

with open(path) as f:
    first = f.readline().strip()
    rest = [l.strip() for l in f.readlines()]

os.unlink(path)
print(first)
print(rest)
"#;
    assert_eq!(run_python(script), vec!["line1", "['line2', 'line3']"]);
}

#[test]
fn test_write_binary_file() {
    let script = r#"
import tempfile, os

with tempfile.NamedTemporaryFile(delete=False) as f:
    path = f.name
    f.write(b'\x00\x01\x02\x03')

with open(path, 'rb') as f:
    data = f.read()

os.unlink(path)
print(list(data))
"#;
    assert_eq!(run_python(script), vec!["[0, 1, 2, 3]"]);
}

#[test]
fn test_seek_tell() {
    let script = r#"
import tempfile, os

with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
    path = f.name
    f.write("abcdef")

with open(path, 'r') as f:
    f.seek(3)
    print(f.tell())
    print(f.read(2))

os.unlink(path)
"#;
    assert_eq!(run_python(script), vec!["3", "de"]);
}

#[test]
fn test_append_mode() {
    let script = r#"
import tempfile, os

with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
    path = f.name
    f.write("first\n")

with open(path, 'a') as f:
    f.write("second\n")

with open(path) as f:
    lines = [l.strip() for l in f.readlines()]

os.unlink(path)
print(lines)
"#;
    assert_eq!(run_python(script), vec!["['first', 'second']"]);
}

#[test]
fn test_file_iteration() {
    let script = r#"
import tempfile, os

with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
    path = f.name
    f.write("a\nb\nc\n")

lines = []
with open(path) as f:
    for line in f:
        lines.append(line.strip())

os.unlink(path)
print(lines)
"#;
    assert_eq!(run_python(script), vec!["['a', 'b', 'c']"]);
}
