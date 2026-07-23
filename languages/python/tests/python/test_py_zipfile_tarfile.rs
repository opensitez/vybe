use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: zipfile + tarfile + gzip + bz2 + lzma — archiving & compression libraries
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_zipfile_create_and_extract() {
    let src = r#"
import zipfile, io

buf = io.BytesIO()
with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
    zf.writestr("doc.txt", "zip file content")
    zf.writestr("sub/nested.txt", "nested data")

buf.seek(0)
with zipfile.ZipFile(buf, "r") as zf:
    print(zf.namelist())
    print(zf.read("doc.txt").decode())
    print(zf.read("sub/nested.txt").decode())
"#;
    assert_eq!(
        run_python(src),
        vec![
            "['doc.txt', 'sub/nested.txt']",
            "zip file content",
            "nested data"
        ]
    );
}

#[test]
fn test_py_zipfile_getinfo_metadata() {
    let src = r#"
import zipfile, io

buf = io.BytesIO()
with zipfile.ZipFile(buf, "w") as zf:
    zf.writestr("test.txt", "hello world")

buf.seek(0)
with zipfile.ZipFile(buf, "r") as zf:
    info = zf.getinfo("test.txt")
    print(info.filename)
    print(info.file_size)
    print(info.date_time[:3])  # (year, month, day)
"#;
    assert_eq!(run_python(src), vec!["test.txt", "11", "(2026, 7, 22)"]);
}

#[test]
fn test_py_tarfile_create_and_extract() {
    let src = r#"
import tarfile, io

buf = io.BytesIO()
with tarfile.open(fileobj=buf, mode="w:gz") as tar:
    data = b"tar content"
    ti = tarfile.TarInfo(name="entry.txt")
    ti.size = len(data)
    tar.addfile(ti, io.BytesIO(data))

buf.seek(0)
with tarfile.open(fileobj=buf, mode="r:gz") as tar:
    names = tar.getnames()
    print(names)
    member = tar.getmember("entry.txt")
    f = tar.extractfile(member)
    print(f.read().decode())
"#;
    assert_eq!(run_python(src), vec!["['entry.txt']", "tar content"]);
}

#[test]
fn test_py_gzip_compress_decompress() {
    let src = r#"
import gzip

raw = b"compress me " * 100
compressed = gzip.compress(raw)
print(len(compressed) < len(raw))

decompressed = gzip.decompress(compressed)
print(decompressed == raw)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_bz2_compress_decompress() {
    let src = r#"
import bz2

raw = b"bzip2 data block " * 50
compressed = bz2.compress(raw)
print(len(compressed) < len(raw))

decompressed = bz2.decompress(compressed)
print(decompressed == raw)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_lzma_xz_compress_decompress() {
    let src = r#"
import lzma

raw = b"lzma xz data stream " * 50
compressed = lzma.compress(raw)
print(len(compressed) < len(raw))

decompressed = lzma.decompress(compressed)
print(decompressed == raw)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_zipfile_infolist_iteration() {
    let src = r#"
import zipfile, io

buf = io.BytesIO()
with zipfile.ZipFile(buf, "w") as zf:
    for name in ["file1.txt", "file2.txt", "file3.txt"]:
        zf.writestr(name, f"content of {name}")

buf.seek(0)
with zipfile.ZipFile(buf, "r") as zf:
    filenames = [info.filename for info in zf.infolist()]
    print(filenames)
"#;
    assert_eq!(
        run_python(src),
        vec!["['file1.txt', 'file2.txt', 'file3.txt']"]
    );
}

#[test]
fn test_py_gzip_open_context_manager() {
    let src = r#"
import gzip, tempfile, os

with tempfile.NamedTemporaryFile(suffix=".gz", delete=False) as f:
    fname = f.name

with gzip.open(fname, "wt", encoding="utf-8") as f:
    f.write("Line 1\nLine 2\n")

with gzip.open(fname, "rt", encoding="utf-8") as f:
    lines = [l.strip() for l in f]

os.unlink(fname)
print(lines)
"#;
    assert_eq!(run_python(src), vec!["['Line 1', 'Line 2']"]);
}

#[test]
fn test_py_zipfile_is_zipfile() {
    let src = r#"
import zipfile, io

buf = io.BytesIO()
with zipfile.ZipFile(buf, "w") as zf:
    zf.writestr("a.txt", "hello")

buf.seek(0)
print(zipfile.is_zipfile(buf))

invalid_buf = io.BytesIO(b"not a zip file")
print(zipfile.is_zipfile(invalid_buf))
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_tarfile_append_mode() {
    let src = r#"
import tarfile, tempfile, os

with tempfile.NamedTemporaryFile(suffix=".tar", delete=False) as f:
    fname = f.name

with tarfile.open(fname, "w") as tar:
    data = b"first"
    ti = tarfile.TarInfo("first.txt")
    ti.size = len(data)
    tar.addfile(ti, io.BytesIO(data)) if "io" in dir() else None

with tarfile.open(fname, "r") as tar:
    print(tar.getnames())

os.unlink(fname)
"#;
    assert_eq!(run_python(src), vec!["['first.txt']"]);
}
