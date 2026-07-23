use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: shutil + tempfile — file copy, tree operations, disk usage, archiving, temporary directories
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_shutil_copy_file() {
    let src = r#"
import shutil, tempfile, os

with tempfile.TemporaryDirectory() as tmp:
    src = os.path.join(tmp, "src.txt")
    dst = os.path.join(tmp, "dst.txt")
    with open(src, "w") as f:
        f.write("Hello Copy")

    shutil.copy(src, dst)
    print(os.path.exists(dst))
    with open(dst) as f:
        print(f.read())
"#;
    assert_eq!(run_python(src), vec!["True", "Hello Copy"]);
}

#[test]
fn test_py_shutil_copytree_rmtree() {
    let src = r#"
import shutil, tempfile, os

with tempfile.TemporaryDirectory() as tmp:
    src_dir = os.path.join(tmp, "src")
    os.makedirs(os.path.join(src_dir, "sub"))
    with open(os.path.join(src_dir, "sub", "file.txt"), "w") as f:
        f.write("nested content")

    dst_dir = os.path.join(tmp, "dst")
    shutil.copytree(src_dir, dst_dir)
    print(os.path.exists(os.path.join(dst_dir, "sub", "file.txt")))

    shutil.rmtree(dst_dir)
    print(os.path.exists(dst_dir))
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_shutil_move_file() {
    let src = r#"
import shutil, tempfile, os

with tempfile.TemporaryDirectory() as tmp:
    src = os.path.join(tmp, "original.txt")
    dst = os.path.join(tmp, "moved.txt")
    with open(src, "w") as f:
        f.write("movable")

    shutil.move(src, dst)
    print(os.path.exists(src))
    print(os.path.exists(dst))
"#;
    assert_eq!(run_python(src), vec!["False", "True"]);
}

#[test]
fn test_py_shutil_disk_usage() {
    let src = r#"
import shutil

usage = shutil.disk_usage(".")
print(usage.total > 0)
print(usage.used > 0)
print(usage.free > 0)
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_shutil_which_executable() {
    let src = r#"
import shutil

python_path = shutil.which("python3") or shutil.which("python")
print(python_path is not None)
print(shutil.which("nonexistent_binary_xyz_123") is None)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_shutil_make_archive_unpack_archive() {
    let src = r#"
import shutil, tempfile, os

with tempfile.TemporaryDirectory() as tmp:
    data_dir = os.path.join(tmp, "data")
    os.makedirs(data_dir)
    with open(os.path.join(data_dir, "doc.txt"), "w") as f:
        f.write("archive document")

    archive_base = os.path.join(tmp, "my_archive")
    archive_path = shutil.make_archive(archive_base, "zip", data_dir)
    print(os.path.exists(archive_path))

    extract_dir = os.path.join(tmp, "extracted")
    shutil.unpack_archive(archive_path, extract_dir)
    print(os.path.exists(os.path.join(extract_dir, "doc.txt")))
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_tempfile_spooled_temporary_file() {
    let src = r#"
import tempfile

with tempfile.SpooledTemporaryFile(max_size=100, mode="w+") as f:
    f.write("small data")
    f.seek(0)
    print(f.read())
    print(f._rolled)  # should not be rolled to disk yet

    f.write("x" * 200)
    f.seek(0)
    print(len(f.read()))
    print(f._rolled)  # now rolled to disk
"#;
    assert_eq!(run_python(src), vec!["small data", "False", "210", "True"]);
}

#[test]
fn test_py_shutil_chown_copymode() {
    let src = r#"
import shutil, tempfile, os

with tempfile.TemporaryDirectory() as tmp:
    f1 = os.path.join(tmp, "f1.txt")
    f2 = os.path.join(tmp, "f2.txt")
    open(f1, "w").close()
    open(f2, "w").close()

    os.chmod(f1, 0o755)
    shutil.copymode(f1, f2)
    print(oct(os.stat(f2).st_mode & 0o777))
"#;
    assert_eq!(run_python(src), vec!["0o755"]);
}

#[test]
fn test_py_tempfile_gettempdir_gettempprefix() {
    let src = r#"
import tempfile

print(isinstance(tempfile.gettempdir(), str))
print(isinstance(tempfile.gettempprefix(), str))
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_shutil_get_archive_formats() {
    let src = r#"
import shutil

formats = [name for name, _ in shutil.get_archive_formats()]
print("zip" in formats)
print("tar" in formats)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}
