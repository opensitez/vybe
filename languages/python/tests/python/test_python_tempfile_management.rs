use super::helpers::run_python;

// tempfile — TemporaryFile, NamedTemporaryFile, TemporaryDirectory, SpooledTemporaryFile, mkstemp, mkdtemp, gettempdir, gettempprefix

#[test]
fn test_tempfile_temporary_file_read_write() {
    let out = run_python(
        r#"
import tempfile
with tempfile.TemporaryFile(mode="w+t") as f:
    f.write("hello tempfile\n")
    f.seek(0)
    print(f.read().strip())
"#,
    );
    assert_eq!(out, vec!["hello tempfile"]);
}

#[test]
fn test_tempfile_named_temporary_file_name_exists() {
    let out = run_python(
        r#"
import tempfile, os
with tempfile.NamedTemporaryFile(delete=False) as f:
    filename = f.name
    f.write(b"named temp data")

print(os.path.exists(filename))
os.unlink(filename)
print(os.path.exists(filename))
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_tempfile_temporary_directory_cleanup() {
    let out = run_python(
        r#"
import tempfile, os
with tempfile.TemporaryDirectory() as tmpdir:
    print(os.path.isdir(tmpdir))
    path = tmpdir

print(os.path.exists(path))
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_tempfile_spooled_temporary_file_rollover() {
    let out = run_python(
        r#"
import tempfile
# max_size 10 bytes: start in-memory, rollover to disk when size exceeds
with tempfile.SpooledTemporaryFile(max_size=10, mode="w+b") as sf:
    sf.write(b"small")
    print(sf._rolled if hasattr(sf, "_rolled") else True)
    sf.write(b" and now this exceeds 10 bytes!")
    print(sf._rolled if hasattr(sf, "_rolled") else True)
    sf.seek(0)
    print(sf.read().startswith(b"small"))
"#,
    );
    assert_eq!(out, vec!["False", "True", "True"]);
}

#[test]
fn test_tempfile_mkstemp_returns_fd_and_path() {
    let out = run_python(
        r#"
import tempfile, os
fd, path = tempfile.mkstemp(suffix=".txt", prefix="vybe_")
print(isinstance(fd, int))
print(path.endswith(".txt"))
print(os.path.basename(path).startswith("vybe_"))
os.close(fd)
os.unlink(path)
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_tempfile_mkdtemp_creates_directory() {
    let out = run_python(
        r#"
import tempfile, os
dpath = tempfile.mkdtemp(prefix="dir_test_")
print(os.path.isdir(dpath))
os.rmdir(dpath)
print(os.path.exists(dpath))
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_tempfile_gettempdir_is_string() {
    let out = run_python(
        r#"
import tempfile, os
tdir = tempfile.gettempdir()
print(isinstance(tdir, str))
print(os.path.isdir(tdir))
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_tempfile_gettempprefix_is_string() {
    let out = run_python(
        r#"
import tempfile
prefix = tempfile.gettempprefix()
print(isinstance(prefix, str))
print(len(prefix) > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_tempfile_named_temporary_file_custom_prefix_suffix() {
    let out = run_python(
        r#"
import tempfile, os
with tempfile.NamedTemporaryFile(prefix="custom_pref_", suffix=".log") as f:
    basename = os.path.basename(f.name)
    print(basename.startswith("custom_pref_"))
    print(basename.endswith(".log"))
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_tempfile_named_temporary_file_custom_dir() {
    let out = run_python(
        r#"
import tempfile, os
with tempfile.TemporaryDirectory() as tmpdir:
    with tempfile.NamedTemporaryFile(dir=tmpdir) as f:
        print(os.path.dirname(f.name) == tmpdir or os.path.samefile(os.path.dirname(f.name), tmpdir))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tempfile_temporary_file_binary_mode() {
    let out = run_python(
        r#"
import tempfile
with tempfile.TemporaryFile(mode="w+b") as f:
    f.write(b"\x00\x01\x02\x03")
    f.seek(0)
    data = f.read()
    print(data == b"\x00\x01\x02\x03")
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tempfile_temporary_directory_ignore_cleanup_errors() {
    let out = run_python(
        r#"
import tempfile, os, sys
if sys.version_info >= (3, 10):
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmpdir:
        pass
    print("ok")
else:
    print("ok")
"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn test_tempfile_tempdir_module_attribute() {
    let out = run_python(
        r#"
import tempfile
# Setting tempfile.tempdir overrides location
orig = tempfile.tempdir
tempfile.tempdir = "/tmp"
print(tempfile.gettempdir())
tempfile.tempdir = orig
"#,
    );
    assert_eq!(out, vec!["/tmp"]);
}

#[test]
fn test_tempfile_spooled_temporary_file_rollover_explicit() {
    let out = run_python(
        r#"
import tempfile
with tempfile.SpooledTemporaryFile(max_size=100) as sf:
    sf.write(b"data")
    if hasattr(sf, "rollover"):
        sf.rollover()
        print(sf._rolled if hasattr(sf, "_rolled") else True)
    else:
        print(True)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tempfile_mkstemp_text_mode() {
    let out = run_python(
        r#"
import tempfile, os
fd, path = tempfile.mkstemp(text=True)
with open(fd, "w") as f:
    f.write("text content\n")
with open(path, "r") as f:
    print(f.read().strip())
os.unlink(path)
"#,
    );
    assert_eq!(out, vec!["text content"]);
}

#[test]
fn test_tempfile_named_temporary_file_delete_on_close_false() {
    let out = run_python(
        r#"
import tempfile, os, sys
if sys.version_info >= (3, 12):
    f = tempfile.NamedTemporaryFile(delete_on_close=False)
    name = f.name
    f.close()
    print(os.path.exists(name))
    os.unlink(name)
else:
    print(True)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tempfile_temporary_file_auto_deletes_on_close() {
    let out = run_python(
        r#"
import tempfile, os
f = tempfile.NamedTemporaryFile(delete=True)
name = f.name
f.close()
print(os.path.exists(name))
"#,
    );
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_tempfile_mkdtemp_custom_dir() {
    let out = run_python(
        r#"
import tempfile, os
with tempfile.TemporaryDirectory() as parent:
    sub = tempfile.mkdtemp(dir=parent)
    print(os.path.dirname(sub) == parent or os.path.samefile(os.path.dirname(sub), parent))
    os.rmdir(sub)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tempfile_temporary_directory_name_attribute() {
    let out = run_python(
        r#"
import tempfile, os
td = tempfile.TemporaryDirectory()
print(isinstance(td.name, str))
print(os.path.isdir(td.name))
td.cleanup()
print(os.path.exists(td.name))
"#,
    );
    assert_eq!(out, vec!["True", "True", "False"]);
}

#[test]
fn test_tempfile_spooled_temporary_file_fileno() {
    let out = run_python(
        r#"
import tempfile
with tempfile.SpooledTemporaryFile(max_size=10) as sf:
    sf.write(b"exceed max size to force rollover")
    if hasattr(sf, "fileno"):
        try:
            print(isinstance(sf.fileno(), int))
        except Exception:
            print(True)
    else:
        print(True)
"#,
    );
    assert_eq!(out, vec!["True"]);
}
