use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: OS & Pathlib Filesystem Traversal — Path methods, glob, rglob, read/write, os.walk, os.scandir, os.stat
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_pathlib_path_join_resolve_name_stem_suffix() {
    let src = r#"
from pathlib import Path

p = Path("/usr/local/bin/python.tar.gz")
print(p.name)
print(p.stem)
print(p.suffix)
print(p.suffixes)
print(p.parent.name)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "python.tar.gz",
            "python.tar",
            ".gz",
            "['.tar', '.gz']",
            "bin"
        ]
    );
}

#[test]
fn test_py_pathlib_read_write_text_bytes() {
    let src = r#"
import tempfile
from pathlib import Path

with tempfile.TemporaryDirectory() as tmpdir:
    p = Path(tmpdir) / "file.txt"
    p.write_text("Hello Pathlib", encoding="utf-8")
    print(p.exists())
    print(p.read_text(encoding="utf-8"))
    
    b_file = Path(tmpdir) / "binary.bin"
    b_file.write_bytes(b"\x00\x01\x02")
    print(b_file.read_bytes())
"#;
    assert_eq!(
        run_python(src),
        vec!["True", "Hello Pathlib", "b'\\x00\\x01\\x02'"]
    );
}

#[test]
fn test_py_pathlib_mkdir_rmdir_parents() {
    let src = r#"
import tempfile
from pathlib import Path

with tempfile.TemporaryDirectory() as tmpdir:
    nested = Path(tmpdir) / "a" / "b" / "c"
    nested.mkdir(parents=True, exist_ok=True)
    print(nested.is_dir())
    print(nested.is_file())
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_pathlib_glob_rglob_pattern_matching() {
    let src = r#"
import tempfile
from pathlib import Path

with tempfile.TemporaryDirectory() as tmpdir:
    root = Path(tmpdir)
    (root / "a.py").write_text("code_a")
    (root / "b.txt").write_text("text_b")
    sub = root / "sub"
    sub.mkdir()
    (sub / "c.py").write_text("code_c")

    py_files = sorted([p.name for p in root.rglob("*.py")])
    print(py_files)
"#;
    assert_eq!(run_python(src), vec!["['a.py', 'c.py']"]);
}

#[test]
fn test_py_os_walk_tree_traversal() {
    let src = r#"
import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    sub = os.path.join(tmpdir, "subdir")
    os.mkdir(sub)
    open(os.path.join(tmpdir, "top.txt"), "w").close()
    open(os.path.join(sub, "sub.txt"), "w").close()

    visited_files = []
    for root, dirs, files in os.walk(tmpdir):
        visited_files.extend(files)

    print(sorted(visited_files))
"#;
    assert_eq!(run_python(src), vec!["['sub.txt', 'top.txt']"]);
}

#[test]
fn test_py_os_scandir_entry_attributes() {
    let src = r#"
import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    fpath = os.path.join(tmpdir, "test.txt")
    with open(fpath, "w") as f:
        f.write("content")

    with os.scandir(tmpdir) as entries:
        for entry in entries:
            print(entry.name)
            print(entry.is_file())
            print(entry.stat().st_size > 0)
"#;
    assert_eq!(run_python(src), vec!["test.txt", "True", "True"]);
}

#[test]
fn test_py_pathlib_with_suffix_with_name() {
    let src = r#"
from pathlib import Path

p = Path("script.py")
print(p.with_suffix(".txt").name)
print(p.with_name("main.py").name)
"#;
    assert_eq!(run_python(src), vec!["script.txt", "main.py"]);
}

#[test]
fn test_py_os_path_split_splitext_basename_dirname() {
    let src = r#"
import os

path = "/usr/local/bin/python.py"
print(os.path.basename(path))
print(os.path.dirname(path))
print(os.path.splitext(path))
print(os.path.split(path))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "python.py",
            "/usr/local/bin",
            "('/usr/local/bin/python', '.py')",
            "('/usr/local/bin', 'python.py')"
        ]
    );
}

#[test]
fn test_py_pathlib_relative_to() {
    let src = r#"
from pathlib import Path

p = Path("/var/log/nginx/access.log")
rel = p.relative_to("/var/log")
print(rel)
"#;
    assert_eq!(run_python(src), vec!["nginx/access.log"]);
}

#[test]
fn test_py_os_stat_result_attributes() {
    let src = r#"
import os, tempfile

with tempfile.NamedTemporaryFile(delete=False) as f:
    f.write(b"12345")
    fname = f.name

st = os.stat(fname)
print(st.st_size)
print(isinstance(st.st_mtime, float))
os.unlink(fname)
"#;
    assert_eq!(run_python(src), vec!["5", "True"]);
}
