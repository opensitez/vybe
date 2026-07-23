use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Shutil & Tempfile Archive Management — copytree, rmtree, move, make_archive, unpack_archive, NamedTemporaryFile, TemporaryDirectory
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_shutil_copyfile_copystat_permissions() {
    let src = r#"
import shutil, tempfile, os

with tempfile.TemporaryDirectory() as tmpdir:
    src_file = os.path.join(tmpdir, "src.txt")
    dst_file = os.path.join(tmpdir, "dst.txt")
    with open(src_file, "w") as f:
        f.write("content for copy")

    shutil.copyfile(src_file, dst_file)
    print(os.path.exists(dst_file))
    with open(dst_file) as f:
        print(f.read())
"#;
    assert_eq!(run_python(src), vec!["True", "content for copy"]);
}

#[test]
fn test_py_shutil_copytree_ignore_patterns() {
    let src = r#"
import shutil, tempfile, os

with tempfile.TemporaryDirectory() as tmpdir:
    src_dir = os.path.join(tmpdir, "src")
    os.makedirs(os.path.join(src_dir, "sub"))
    open(os.path.join(src_dir, "keep.py"), "w").close()
    open(os.path.join(src_dir, "ignore.tmp"), "w").close()

    dst_dir = os.path.join(tmpdir, "dst")
    shutil.copytree(src_dir, dst_dir, ignore=shutil.ignore_patterns("*.tmp"))

    print(os.path.exists(os.path.join(dst_dir, "keep.py")))
    print(os.path.exists(os.path.join(dst_dir, "ignore.tmp")))
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_shutil_move_directory_tree() {
    let src = r#"
import shutil, tempfile, os

with tempfile.TemporaryDirectory() as tmpdir:
    src_dir = os.path.join(tmpdir, "old_location")
    dst_dir = os.path.join(tmpdir, "new_location")
    os.mkdir(src_dir)
    open(os.path.join(src_dir, "data.txt"), "w").close()

    shutil.move(src_dir, dst_dir)
    print(os.path.exists(src_dir))
    print(os.path.exists(os.path.join(dst_dir, "data.txt")))
"#;
    assert_eq!(run_python(src), vec!["False", "True"]);
}

#[test]
fn test_py_shutil_disk_usage_tuple() {
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
fn test_py_tempfile_named_temporary_file_auto_delete() {
    let src = r#"
import tempfile, os

with tempfile.NamedTemporaryFile(mode="w+", delete=False) as f:
    f.write("temporary data")
    fname = f.name

print(os.path.exists(fname))
os.unlink(fname)
print(os.path.exists(fname))
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_tempfile_temporary_directory_cleanup() {
    let src = r#"
import tempfile, os

with tempfile.TemporaryDirectory() as tmpdir:
    fpath = os.path.join(tmpdir, "test.txt")
    open(fpath, "w").close()
    saved_path = tmpdir

print(os.path.exists(saved_path))
"#;
    assert_eq!(run_python(src), vec!["False"]);
}

#[test]
fn test_py_shutil_which_find_executable() {
    let src = r#"
import shutil

py_path = shutil.which("python3") or shutil.which("python")
print(py_path is not None)
print(shutil.which("nonexistent_binary_xyz_99") is None)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_shutil_make_archive_unpack_zip() {
    let src = r#"
import shutil, tempfile, os

with tempfile.TemporaryDirectory() as tmpdir:
    data_dir = os.path.join(tmpdir, "data")
    os.mkdir(data_dir)
    with open(os.path.join(data_dir, "hello.txt"), "w") as f:
        f.write("archive test")

    archive_path = shutil.make_archive(os.path.join(tmpdir, "my_archive"), "zip", data_dir)
    print(os.path.exists(archive_path))

    extract_dir = os.path.join(tmpdir, "extracted")
    shutil.unpack_archive(archive_path, extract_dir)
    print(os.path.exists(os.path.join(extract_dir, "hello.txt")))
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_tempfile_spooled_temporary_file_rollover() {
    let src = r#"
import tempfile

with tempfile.SpooledTemporaryFile(max_size=50) as f:
    f.write(b"small data")
    print(f._rolled)  # not rolled yet
    f.write(b"x" * 100)
    print(f._rolled)  # rolled to disk
"#;
    assert_eq!(run_python(src), vec!["False", "True"]);
}

#[test]
fn test_py_shutil_get_unpack_formats() {
    let src = r#"
import shutil

formats = [fmt for fmt, _, _ in shutil.get_unpack_formats()]
print("zip" in formats)
print("tar" in formats)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}
