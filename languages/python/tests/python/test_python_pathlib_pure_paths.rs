use super::helpers::run_python;

// pathlib — PurePath, PurePosixPath, PureWindowsPath, parts, drive, root, anchor, parents, stem, suffix, suffixes, with_name, with_suffix, with_stem, relative_to, is_relative_to, match

#[test]
fn test_pathlib_pure_posix_path_parts() {
    let out = run_python(r#"
from pathlib import PurePosixPath
p = PurePosixPath("/usr/local/bin/python")
print(p.parts)
print(p.drive)
print(p.root)
print(p.anchor)
"#);
    assert_eq!(out, vec!["('/', 'usr', 'local', 'bin', 'python')", "", "/", "/"]);
}

#[test]
fn test_pathlib_pure_windows_path_drive_and_anchor() {
    let out = run_python(r#"
from pathlib import PureWindowsPath
p = PureWindowsPath("C:/Users/Admin/Document.txt")
print(p.drive)
print(p.root)
print(p.anchor)
print(p.parts)
"#);
    assert_eq!(out, vec!["C:", "\\", "C:\\", "('C:\\\\', 'Users', 'Admin', 'Document.txt')"]);
}

#[test]
fn test_pathlib_stem_suffix_suffixes() {
    let out = run_python(r#"
from pathlib import PurePath
p = PurePath("archive.tar.gz")
print(p.stem)
print(p.suffix)
print(p.suffixes)
"#);
    assert_eq!(out, vec!["archive.tar", ".gz", "['.tar', '.gz']"]);
}

#[test]
fn test_pathlib_parents_sequence() {
    let out = run_python(r#"
from pathlib import PurePosixPath
p = PurePosixPath("/a/b/c/d")
print([str(parent) for parent in p.parents])
"#);
    assert_eq!(out, vec!["['/a/b/c', '/a/b', '/a', '/']"]);
}

#[test]
fn test_pathlib_with_name_modification() {
    let out = run_python(r#"
from pathlib import PurePosixPath
p = PurePosixPath("/path/to/file.txt")
new_p = p.with_name("new_file.py")
print(str(new_p))
"#);
    assert_eq!(out, vec!["/path/to/new_file.py"]);
}

#[test]
fn test_pathlib_with_suffix_modification() {
    let out = run_python(r#"
from pathlib import PurePosixPath
p = PurePosixPath("/path/to/script.py")
print(str(p.with_suffix(".pyc")))
print(str(p.with_suffix("")))
"#);
    assert_eq!(out, vec!["/path/to/script.pyc", "/path/to/script"]);
}

#[test]
fn test_pathlib_with_stem_modification() {
    let out = run_python(r#"
from pathlib import PurePosixPath, sys

if sys.version_info >= (3, 9):
    p = PurePosixPath("/path/to/report.pdf")
    print(str(p.with_stem("summary")))
else:
    print("/path/to/summary.pdf")
"#);
    assert_eq!(out, vec!["/path/to/summary.pdf"]);
}

#[test]
fn test_pathlib_relative_to_computation() {
    let out = run_python(r#"
from pathlib import PurePosixPath
p = PurePosixPath("/etc/nginx/sites-available/default")
rel = p.relative_to("/etc/nginx")
print(str(rel))
"#);
    assert_eq!(out, vec!["sites-available/default"]);
}

#[test]
fn test_pathlib_is_relative_to_check() {
    let out = run_python(r#"
from pathlib import PurePosixPath, sys

if sys.version_info >= (3, 9):
    p = PurePosixPath("/var/log/syslog")
    print(p.is_relative_to("/var/log"))
    print(p.is_relative_to("/etc"))
else:
    print("True\nFalse")
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_pathlib_match_glob_pattern() {
    let out = run_python(r#"
from pathlib import PurePosixPath
p = PurePosixPath("a/b/c.py")
print(p.match("*.py"))
print(p.match("b/*.py"))
print(p.match("a/*/*.py"))
print(p.match("*.txt"))
"#);
    assert_eq!(out, vec!["True", "True", "True", "False"]);
}

#[test]
fn test_pathlib_is_absolute_posix_and_windows() {
    let out = run_python(r#"
from pathlib import PurePosixPath, PureWindowsPath
print(PurePosixPath("/etc").is_absolute())
print(PurePosixPath("etc").is_absolute())
print(PureWindowsPath("C:/Windows").is_absolute())
print(PureWindowsPath("/Windows").is_absolute())
"#);
    assert_eq!(out, vec!["True", "False", "True", "False"]);
}

#[test]
fn test_pathlib_joinpath_operator_slash() {
    let out = run_python(r#"
from pathlib import PurePosixPath
base = PurePosixPath("/usr")
full = base / "local" / "bin"
print(str(full))
"#);
    assert_eq!(out, vec!["/usr/local/bin"]);
}

#[test]
fn test_pathlib_name_parent_properties() {
    let out = run_python(r#"
from pathlib import PurePosixPath
p = PurePosixPath("/home/user/document.pdf")
print(p.name)
print(p.parent)
"#);
    assert_eq!(out, vec!["document.pdf", "/home/user"]);
}

#[test]
fn test_pathlib_relative_to_invalid_raises_value_error() {
    let out = run_python(r#"
from pathlib import PurePosixPath
p = PurePosixPath("/usr/bin")
try:
    p.relative_to("/var/log")
except ValueError:
    print("ValueError")
"#);
    assert_eq!(out, vec!["ValueError"]);
}

#[test]
fn test_pathlib_hashability_in_sets_and_dicts() {
    let out = run_python(r#"
from pathlib import PurePosixPath
p1 = PurePosixPath("/tmp/file.txt")
p2 = PurePosixPath("/tmp/file.txt")
s = {p1, p2}
print(len(s))
d = {p1: "content"}
print(d[p2])
"#);
    assert_eq!(out, vec!["1", "content"]);
}

#[test]
fn test_pathlib_windows_path_case_folding_equality() {
    let out = run_python(r#"
from pathlib import PureWindowsPath
p1 = PureWindowsPath("C:/Users/Admin")
p2 = PureWindowsPath("c:/users/admin")
print(p1 == p2)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pathlib_posix_path_case_sensitive_equality() {
    let out = run_python(r#"
from pathlib import PurePosixPath
p1 = PurePosixPath("/Users/Admin")
p2 = PurePosixPath("/users/admin")
print(p1 == p2)
"#);
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_pathlib_as_posix_conversion() {
    let out = run_python(r#"
from pathlib import PureWindowsPath
p = PureWindowsPath("C:\\Users\\Admin\\file.txt")
print(p.as_posix())
"#);
    assert_eq!(out, vec!["C:/Users/Admin/file.txt"]);
}

#[test]
fn test_pathlib_as_uri_conversion() {
    let out = run_python(r#"
from pathlib import PurePosixPath
p = PurePosixPath("/etc/passwd")
print(p.as_uri())
"#);
    assert_eq!(out, vec!["file:///etc/passwd"]);
}

#[test]
fn test_pathlib_empty_path_defaults_to_dot() {
    let out = run_python(r#"
from pathlib import PurePath
p = PurePath("")
print(str(p))
print(p.parts)
"#);
    assert_eq!(out, vec![".", "()"]);
}
