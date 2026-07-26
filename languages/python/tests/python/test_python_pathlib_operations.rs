// Python pathlib operations — Path creation, joining, stem, suffix, parts
use super::helpers::run_python;

#[test]
fn test_pathlib_creation() {
    let script = r#"
from pathlib import Path
p = Path("/usr/local/bin/python")
print(p.name)
print(p.stem)
print(p.suffix)
"#;
    assert_eq!(run_python(script), vec!["python", "python", ""]);
}

#[test]
fn test_pathlib_parent_parts() {
    let script = r#"
from pathlib import Path
p = Path("/usr/local/bin/python")
print(str(p.parent))
print(p.parts[0])
print(p.parts[1])
"#;
    assert_eq!(run_python(script), vec!["/usr/local/bin", "/", "usr"]);
}

#[test]
fn test_pathlib_joinpath() {
    let script = r#"
from pathlib import Path
base = Path("/tmp")
full = base / "subdir" / "file.txt"
print(full.name)
print(full.suffix)
print(str(full.parent))
"#;
    assert_eq!(run_python(script), vec!["file.txt", ".txt", "/tmp/subdir"]);
}

#[test]
fn test_pathlib_with_suffix() {
    let script = r#"
from pathlib import Path
p = Path("document.txt")
p2 = p.with_suffix(".pdf")
print(p2.name)
"#;
    assert_eq!(run_python(script), vec!["document.pdf"]);
}

#[test]
fn test_pathlib_with_name() {
    let script = r#"
from pathlib import Path
p = Path("/home/user/old.txt")
p2 = p.with_name("new.md")
print(p2.name)
print(str(p2.parent))
"#;
    assert_eq!(run_python(script), vec!["new.md", "/home/user"]);
}

#[test]
fn test_pathlib_is_absolute() {
    let script = r#"
from pathlib import Path
print(Path("/absolute/path").is_absolute())
print(Path("relative/path").is_absolute())
"#;
    assert_eq!(run_python(script), vec!["True", "False"]);
}

#[test]
fn test_pathlib_posix_string() {
    let script = r#"
from pathlib import PurePosixPath
p = PurePosixPath("/a/b/c.txt")
print(p.as_posix())
"#;
    assert_eq!(run_python(script), vec!["/a/b/c.txt"]);
}

#[test]
fn test_pathlib_suffixes() {
    let script = r#"
from pathlib import Path
p = Path("archive.tar.gz")
print(p.suffixes)
print(p.stem)
"#;
    assert_eq!(run_python(script), vec!["['.tar', '.gz']", "archive.tar"]);
}
