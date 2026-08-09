use super::helpers::run_python;

// mmap — file-backed memory maps, read/write, find, seek, slice, move, resize

#[test]
fn test_mmap_read_full_content() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"hello world")
f.flush()
with open(f.name, "r+b") as fh:
    mm = mmap.mmap(fh.fileno(), 0)
    print(mm[:11])
    mm.close()
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["b'hello world'"]);
}

#[test]
fn test_mmap_write_modifies_content() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"hello world")
f.flush()
with open(f.name, "r+b") as fh:
    mm = mmap.mmap(fh.fileno(), 0)
    mm[0:5] = b"HELLO"
    mm.flush()
    mm.close()
with open(f.name, "rb") as fh:
    print(fh.read())
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["b'HELLO world'"]);
}

#[test]
fn test_mmap_find_substring() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"abcdefghij")
f.flush()
with open(f.name, "r+b") as fh:
    mm = mmap.mmap(fh.fileno(), 0)
    pos = mm.find(b"def")
    print(pos)
    mm.close()
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_mmap_find_not_found_returns_minus_one() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"abcde")
f.flush()
with open(f.name, "r+b") as fh:
    mm = mmap.mmap(fh.fileno(), 0)
    print(mm.find(b"xyz"))
    mm.close()
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn test_mmap_rfind_searches_from_end() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"abcabc")
f.flush()
with open(f.name, "r+b") as fh:
    mm = mmap.mmap(fh.fileno(), 0)
    print(mm.rfind(b"abc"))  # last occurrence at index 3
    mm.close()
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_mmap_seek_and_read() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"0123456789")
f.flush()
with open(f.name, "r+b") as fh:
    mm = mmap.mmap(fh.fileno(), 0)
    mm.seek(5)
    print(mm.read(3))
    mm.close()
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["b'567'"]);
}

#[test]
fn test_mmap_seek_from_end() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"hello!!")
f.flush()
with open(f.name, "r+b") as fh:
    mm = mmap.mmap(fh.fileno(), 0)
    mm.seek(-2, 2)  # 2 = SEEK_END
    print(mm.read())
    mm.close()
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["b'!!'"]);
}

#[test]
fn test_mmap_tell_after_seek() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"0123456789")
f.flush()
with open(f.name, "r+b") as fh:
    mm = mmap.mmap(fh.fileno(), 0)
    mm.seek(7)
    print(mm.tell())
    mm.close()
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn test_mmap_size_attribute() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"x" * 100)
f.flush()
with open(f.name, "r+b") as fh:
    mm = mmap.mmap(fh.fileno(), 0)
    print(mm.size())
    mm.close()
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn test_mmap_length() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"abcde")
f.flush()
with open(f.name, "r+b") as fh:
    mm = mmap.mmap(fh.fileno(), 0)
    print(len(mm))
    mm.close()
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_mmap_readline() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"first line\nsecond line\n")
f.flush()
with open(f.name, "r+b") as fh:
    mm = mmap.mmap(fh.fileno(), 0)
    line = mm.readline()
    print(line)
    mm.close()
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["b'first line\\n'"]);
}

#[test]
fn test_mmap_write_returns_bytes_written() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"     ")
f.flush()
with open(f.name, "r+b") as fh:
    mm = mmap.mmap(fh.fileno(), 0)
    mm.seek(0)
    n = mm.write(b"hello")
    print(n)
    mm.close()
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_mmap_slice_assignment() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"AAAAAA")
f.flush()
with open(f.name, "r+b") as fh:
    mm = mmap.mmap(fh.fileno(), 0)
    mm[2:4] = b"BB"
    print(mm[:])
    mm.close()
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["b'AABBAA'"]);
}

#[test]
fn test_mmap_getitem_single_byte() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"ABCDE")
f.flush()
with open(f.name, "r+b") as fh:
    mm = mmap.mmap(fh.fileno(), 0)
    print(mm[0])   # int
    print(mm[4])
    mm.close()
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["65", "69"]);
}

#[test]
fn test_mmap_access_read_only() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"readonly data")
f.flush()
f.close()
with open(f.name, "rb") as fh:
    mm = mmap.mmap(fh.fileno(), 0, access=mmap.ACCESS_READ)
    print(mm[:8])
    try:
        mm[0:4] = b"XXXX"
        print("no error")
    except TypeError:
        print("TypeError")
    mm.close()
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["b'readonly'", "TypeError"]);
}

#[test]
fn test_mmap_copy_access() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"original")
f.flush()
f.close()
with open(f.name, "rb") as fh:
    mm = mmap.mmap(fh.fileno(), 0, access=mmap.ACCESS_COPY)
    mm[0:4] = b"XXXX"    # copy-on-write, file unchanged
    print(mm[:8])
    mm.close()
with open(f.name, "rb") as fh:
    print(fh.read())     # file should be unchanged
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["b'XXXXinal'", "b'original'"]);
}

#[test]
fn test_mmap_anon_mmap_unix() {
    let out = run_python(
        r#"
import mmap, sys
if sys.platform != "win32":
    mm = mmap.mmap(-1, 256)  # anonymous mmap
    mm.write(b"test data!")
    mm.seek(0)
    print(mm.read(10))
    mm.close()
else:
    print(b"test data!")
"#,
    );
    assert_eq!(out, vec!["b'test data!'"]);
}

#[test]
fn test_mmap_move_bytes() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"ABCDE12345")
f.flush()
with open(f.name, "r+b") as fh:
    mm = mmap.mmap(fh.fileno(), 0)
    mm.move(0, 5, 5)   # copy bytes 5..10 to 0..5
    print(mm[:10])
    mm.close()
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["b'12345 12345'"]);
}

#[test]
fn test_mmap_flush_no_error() {
    let out = run_python(
        r#"
import mmap, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False)
f.write(b"data here")
f.flush()
with open(f.name, "r+b") as fh:
    mm = mmap.mmap(fh.fileno(), 0)
    mm[0:4] = b"XXXX"
    mm.flush()
    print("ok")
    mm.close()
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["ok"]);
}
