use super::helpers::run_python;

// os — scandir, DirEntry (stat, name, path, is_dir, is_file, is_symlink, inode), walk, pipe, dup, dup2, urandom, cpu_count

#[test]
fn test_os_scandir_iterates_direntries() {
    let out = run_python(
        r#"
import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    fpath = os.path.join(tmpdir, "test.txt")
    with open(fpath, "w") as f:
        f.write("hello")

    entries = list(os.scandir(tmpdir))
    print(len(entries))
    entry = entries[0]
    print(entry.name)
    print(entry.is_file())
    print(entry.is_dir())
"#,
    );
    assert_eq!(out, vec!["1", "test.txt", "True", "False"]);
}

#[test]
fn test_os_scandir_entry_stat_metadata() {
    let out = run_python(
        r#"
import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    fpath = os.path.join(tmpdir, "data.bin")
    with open(fpath, "wb") as f:
        f.write(b"12345")

    with os.scandir(tmpdir) as it:
        entry = next(it)
        st = entry.stat()
        print(st.st_size)
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_os_walk_directory_tree() {
    let out = run_python(
        r#"
import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    sub = os.path.join(tmpdir, "subdir")
    os.mkdir(sub)
    with open(os.path.join(sub, "f.txt"), "w") as f: f.write("x")

    found_dirs = []
    found_files = []
    for root, dirs, files in os.walk(tmpdir):
        found_dirs.extend(dirs)
        found_files.extend(files)

    print("subdir" in found_dirs)
    print("f.txt" in found_files)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_os_pipe_read_write() {
    let out = run_python(
        r#"
import os

r, w = os.pipe()
os.write(w, b"pipe test msg")
os.close(w)

data = os.read(r, 100)
os.close(r)
print(data)
"#,
    );
    assert_eq!(out, vec!["b'pipe test msg'"]);
}

#[test]
fn test_os_dup_file_descriptor() {
    let out = run_python(
        r#"
import os, tempfile

with tempfile.TemporaryFile() as f:
    fd = f.fileno()
    dup_fd = os.dup(fd)
    os.write(dup_fd, b"hello dup")
    os.lseek(fd, 0, 0)
    print(os.read(fd, 100))
    os.close(dup_fd)
"#,
    );
    assert_eq!(out, vec!["b'hello dup'"]);
}

#[test]
fn test_os_urandom_returns_cryptographic_bytes() {
    let out = run_python(
        r#"
import os
b1 = os.urandom(16)
b2 = os.urandom(16)
print(len(b1) == 16)
print(isinstance(b1, bytes))
print(b1 != b2)
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_os_cpu_count_returns_positive_int() {
    let out = run_python(
        r#"
import os
cpus = os.cpu_count()
print(isinstance(cpus, int))
print(cpus > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_os_scandir_entry_path_property() {
    let out = run_python(
        r#"
import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    with open(os.path.join(tmpdir, "sample.log"), "w") as f: f.write("log")
    with os.scandir(tmpdir) as it:
        entry = next(it)
        print(os.path.isabs(entry.path))
        print(entry.path.endswith("sample.log"))
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_os_scandir_entry_inode() {
    let out = run_python(
        r#"
import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    with open(os.path.join(tmpdir, "inode_test"), "w") as f: f.write("a")
    with os.scandir(tmpdir) as it:
        entry = next(it)
        ino = entry.inode()
        print(isinstance(ino, int))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_os_walk_topdown_false() {
    let out = run_python(
        r#"
import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    sub = os.path.join(tmpdir, "a", "b")
    os.makedirs(sub)
    roots = [root for root, _, _ in os.walk(tmpdir, topdown=False)]
    print(roots[0].endswith("b"))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_os_fspath_protocol_support() {
    let out = run_python(
        r#"
import os
from pathlib import Path

p = Path("/tmp/test_fspath")
res = os.fspath(p)
print(res)
print(isinstance(res, str))
"#,
    );
    assert_eq!(out, vec!["/tmp/test_fspath", "True"]);
}

#[test]
fn test_os_device_encoding_check() {
    let out = run_python(
        r#"
import os
enc = os.device_encoding(0)
print(enc is None or isinstance(enc, str))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_os_truncate_file_size() {
    let out = run_python(
        r#"
import os, tempfile

with tempfile.NamedTemporaryFile(delete=False) as f:
    f.write(b"1234567890")
    fname = f.name

os.truncate(fname, 5)
with open(fname, "rb") as f:
    print(f.read())
os.remove(fname)
"#,
    );
    assert_eq!(out, vec!["b'12345'"]);
}

#[test]
fn test_os_strerror_error_messages() {
    let out = run_python(
        r#"
import os
msg = os.strerror(2)  # ENOENT (No such file or directory)
print(isinstance(msg, str))
print(len(msg) > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_os_get_terminal_size_fallback() {
    let out = run_python(
        r#"
import os
try:
    ts = os.get_terminal_size(0)
    print(isinstance(ts.columns, int))
except (OSError, ValueError):
    print("OSError")
"#,
    );
    assert_eq!(out, vec!["OSError"]);
}

#[test]
fn test_os_scandir_closes_iterator_on_context_exit() {
    let out = run_python(
        r#"
import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    with open(os.path.join(tmpdir, "f"), "w") as f: f.write("1")
    it = os.scandir(tmpdir)
    with it:
        entry = next(it)
        print(entry.name)
"#,
    );
    assert_eq!(out, vec!["f"]);
}

#[test]
fn test_os_walk_onerror_callback() {
    let out = run_python(
        r#"
import os

errors = []
def handle_err(err):
    errors.append(err)

list(os.walk("/non_existent_dir_123456789", onerror=handle_err))
print(len(errors) > 0)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_os_replace_atomic_file_rename() {
    let out = run_python(
        r#"
import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    src = os.path.join(tmpdir, "src.txt")
    dst = os.path.join(tmpdir, "dst.txt")
    with open(src, "w") as f: f.write("src content")
    os.replace(src, dst)
    print(os.path.exists(src))
    print(os.path.exists(dst))
"#,
    );
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_os_isatty_file_descriptor() {
    let out = run_python(
        r#"
import os, tempfile

with tempfile.TemporaryFile() as f:
    print(os.isatty(f.fileno()))
"#,
    );
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_os_scandir_symlink_check() {
    let out = run_python(
        r#"
import os, tempfile

with tempfile.TemporaryDirectory() as tmpdir:
    target = os.path.join(tmpdir, "target.txt")
    link = os.path.join(tmpdir, "link.txt")
    with open(target, "w") as f: f.write("target")
    try:
        os.symlink(target, link)
        with os.scandir(tmpdir) as it:
            for entry in it:
                if entry.name == "link.txt":
                    print(entry.is_symlink())
    except OSError:
        # Symlinks may require elevated permissions on Windows
        print(True)
"#,
    );
    assert_eq!(out, vec!["True"]);
}
