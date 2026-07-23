use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: os + sys + pathlib — filesystem, environment, process, sys attributes, Path operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_os_environ_access_and_default() {
    let src = r#"
import os

os.environ["MY_TEST_VAR"] = "hello_world"
print(os.environ["MY_TEST_VAR"])
print(os.getenv("MY_TEST_VAR"))
print(os.getenv("NONEXISTENT_VAR", "fallback"))
del os.environ["MY_TEST_VAR"]
print(os.getenv("MY_TEST_VAR") is None)
"#;
    assert_eq!(
        run_python(src),
        vec!["hello_world", "hello_world", "fallback", "True"]
    );
}

#[test]
fn test_py_os_path_join_and_split() {
    let src = r#"
import os

joined = os.path.join("/usr", "local", "bin", "python")
print(joined)
print(os.path.split(joined))
print(os.path.dirname(joined))
print(os.path.basename(joined))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "/usr/local/bin/python",
            "('/usr/local/bin', 'python')",
            "/usr/local/bin",
            "python"
        ]
    );
}

#[test]
fn test_py_os_path_splitext_and_extension() {
    let src = r#"
import os

path = "/home/user/report.tar.gz"
root, ext = os.path.splitext(path)
print(ext)
print(root)
print(os.path.splitext("file.py")[1])
print(os.path.splitext("noext")[1])
"#;
    assert_eq!(
        run_python(src),
        vec![".gz", "/home/user/report.tar", ".py", ""]
    );
}

#[test]
fn test_py_os_path_abspath_normpath() {
    let src = r#"
import os

path = os.path.normpath("/usr/../usr/./local")
print(path)

relative = os.path.normpath("a/b/../c/./d")
print(relative)
"#;
    assert_eq!(run_python(src), vec!["/usr/local", "a/c/d"]);
}

#[test]
fn test_py_os_makedirs_and_remove_temp() {
    let src = r#"
import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    subdir = os.path.join(tmpdir, "a", "b", "c")
    os.makedirs(subdir)
    print(os.path.isdir(subdir))
    os.removedirs(subdir)
    print(os.path.isdir(subdir))
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_os_listdir_and_scandir() {
    let src = r#"
import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    for name in ["a.txt", "b.py", "c.md"]:
        open(os.path.join(tmpdir, name), "w").close()

    names = sorted(os.listdir(tmpdir))
    print(names)

    entries = sorted([e.name for e in os.scandir(tmpdir)])
    print(entries)
"#;
    assert_eq!(
        run_python(src),
        vec!["['a.txt', 'b.py', 'c.md']", "['a.txt', 'b.py', 'c.md']"]
    );
}

#[test]
fn test_py_os_stat_file_metadata() {
    let src = r#"
import os, tempfile

with tempfile.NamedTemporaryFile(delete=False) as f:
    f.write(b"hello world")
    fname = f.name

stat = os.stat(fname)
print(stat.st_size)
os.remove(fname)
"#;
    assert_eq!(run_python(src), vec!["11"]);
}

#[test]
fn test_py_os_walk_directory_tree() {
    let src = r#"
import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    sub = os.path.join(tmpdir, "sub")
    os.makedirs(sub)
    open(os.path.join(tmpdir, "root.txt"), "w").close()
    open(os.path.join(sub, "child.txt"), "w").close()

    files_found = []
    for root, dirs, files in os.walk(tmpdir):
        for f in files:
            files_found.append(f)
    print(sorted(files_found))
"#;
    assert_eq!(run_python(src), vec!["['child.txt', 'root.txt']"]);
}

#[test]
fn test_py_sys_argv_access() {
    let src = r#"
import sys

print(isinstance(sys.argv, list))
print(len(sys.argv) >= 1)
print(isinstance(sys.argv[0], str))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_sys_path_manipulation() {
    let src = r#"
import sys

original_len = len(sys.path)
sys.path.insert(0, "/fake/path")
print(sys.path[0])
print(len(sys.path) == original_len + 1)
sys.path.pop(0)
"#;
    assert_eq!(run_python(src), vec!["/fake/path", "True"]);
}

#[test]
fn test_py_sys_version_and_platform() {
    let src = r#"
import sys

print(isinstance(sys.version, str))
print(isinstance(sys.version_info.major, int))
print(sys.version_info.major >= 3)
print(isinstance(sys.platform, str))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True", "True"]);
}

#[test]
fn test_py_sys_getrefcount() {
    let src = r#"
import sys

x = []
ref1 = sys.getrefcount(x)  # at least 2 (x + argument)
y = x
ref2 = sys.getrefcount(x)
print(ref2 > ref1)  # adding y increases refcount
del y
ref3 = sys.getrefcount(x)
print(ref3 == ref1)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_sys_recursion_limit() {
    let src = r#"
import sys

limit = sys.getrecursionlimit()
print(limit > 0)
sys.setrecursionlimit(500)
print(sys.getrecursionlimit())
sys.setrecursionlimit(limit)  # restore
"#;
    assert_eq!(run_python(src), vec!["True", "500"]);
}

#[test]
fn test_py_pathlib_path_construction() {
    let src = r#"
from pathlib import Path

p = Path("/usr") / "local" / "bin"
print(str(p))
print(p.parts)
print(p.name)
print(p.parent)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "/usr/local/bin",
            "('/', 'usr', 'local', 'bin')",
            "bin",
            "/usr/local"
        ]
    );
}

#[test]
fn test_py_pathlib_stem_suffix_with_suffix() {
    let src = r#"
from pathlib import Path

p = Path("/home/user/report.txt")
print(p.stem)
print(p.suffix)
print(p.with_suffix(".md"))
print(p.with_name("notes.csv"))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "report",
            ".txt",
            "/home/user/report.md",
            "/home/user/notes.csv"
        ]
    );
}

#[test]
fn test_py_pathlib_read_write_text() {
    let src = r#"
import tempfile
from pathlib import Path

with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
    tmp = Path(f.name)

tmp.write_text("Hello, pathlib!")
content = tmp.read_text()
print(content)
tmp.unlink()
"#;
    assert_eq!(run_python(src), vec!["Hello, pathlib!"]);
}

#[test]
fn test_py_pathlib_glob_rglob() {
    let src = r#"
import tempfile, os
from pathlib import Path

with tempfile.TemporaryDirectory() as tmpdir:
    base = Path(tmpdir)
    (base / "a.py").write_text("")
    (base / "b.py").write_text("")
    (base / "c.txt").write_text("")

    py_files = sorted([p.name for p in base.glob("*.py")])
    all_files = sorted([p.name for p in base.rglob("*") if p.is_file()])
    print(py_files)
    print(all_files)
"#;
    assert_eq!(
        run_python(src),
        vec!["['a.py', 'b.py']", "['a.py', 'b.py', 'c.txt']"]
    );
}

#[test]
fn test_py_pathlib_mkdir_and_exists() {
    let src = r#"
import tempfile
from pathlib import Path

with tempfile.TemporaryDirectory() as tmpdir:
    d = Path(tmpdir) / "nested" / "dir"
    d.mkdir(parents=True, exist_ok=True)
    print(d.exists())
    print(d.is_dir())
    d.rmdir()
    d.parent.rmdir()
    print(d.exists())
"#;
    assert_eq!(run_python(src), vec!["True", "True", "False"]);
}

#[test]
fn test_py_os_cpu_count_and_process_id() {
    let src = r#"
import os

print(os.getpid() > 0)
print(os.cpu_count() is None or os.cpu_count() > 0)
print(isinstance(os.getcwd(), str))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}
