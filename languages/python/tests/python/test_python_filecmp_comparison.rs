use super::helpers::run_python;

// filecmp — cmp, cmpfiles, dircmp

#[test]
fn test_filecmp_cmp_identical_files() {
    let out = run_python(
        r#"
import filecmp, tempfile, os
d = tempfile.mkdtemp()
a = os.path.join(d, "a.txt")
b = os.path.join(d, "b.txt")
for p in [a, b]:
    with open(p, "w") as f:
        f.write("same content")
print(filecmp.cmp(a, b, shallow=False))
import shutil; shutil.rmtree(d)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_filecmp_cmp_different_files() {
    let out = run_python(
        r#"
import filecmp, tempfile, os
d = tempfile.mkdtemp()
a = os.path.join(d, "a.txt")
b = os.path.join(d, "b.txt")
with open(a, "w") as f: f.write("aaa")
with open(b, "w") as f: f.write("bbb")
print(filecmp.cmp(a, b, shallow=False))
import shutil; shutil.rmtree(d)
"#,
    );
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_filecmp_cmp_same_file_is_true() {
    let out = run_python(
        r#"
import filecmp, tempfile, os
f = tempfile.NamedTemporaryFile(delete=False, mode="w", suffix=".txt")
f.write("content")
f.close()
print(filecmp.cmp(f.name, f.name))
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_filecmp_cmpfiles_match() {
    let out = run_python(
        r#"
import filecmp, tempfile, os, shutil
d1 = tempfile.mkdtemp()
d2 = tempfile.mkdtemp()
for d in [d1, d2]:
    with open(os.path.join(d, "x.txt"), "w") as f:
        f.write("identical")
match, mismatch, errors = filecmp.cmpfiles(d1, d2, ["x.txt"], shallow=False)
print(match)
print(mismatch)
shutil.rmtree(d1); shutil.rmtree(d2)
"#,
    );
    assert_eq!(out, vec!["['x.txt']", "[]"]);
}

#[test]
fn test_filecmp_cmpfiles_mismatch() {
    let out = run_python(
        r#"
import filecmp, tempfile, os, shutil
d1 = tempfile.mkdtemp()
d2 = tempfile.mkdtemp()
with open(os.path.join(d1, "x.txt"), "w") as f: f.write("aaa")
with open(os.path.join(d2, "x.txt"), "w") as f: f.write("bbb")
match, mismatch, errors = filecmp.cmpfiles(d1, d2, ["x.txt"], shallow=False)
print(match)
print(mismatch)
shutil.rmtree(d1); shutil.rmtree(d2)
"#,
    );
    assert_eq!(out, vec!["[]", "['x.txt']"]);
}

#[test]
fn test_filecmp_cmpfiles_missing_in_right() {
    let out = run_python(
        r#"
import filecmp, tempfile, os, shutil
d1 = tempfile.mkdtemp()
d2 = tempfile.mkdtemp()
with open(os.path.join(d1, "x.txt"), "w") as f: f.write("aaa")
match, mismatch, errors = filecmp.cmpfiles(d1, d2, ["x.txt"], shallow=False)
print(errors)
shutil.rmtree(d1); shutil.rmtree(d2)
"#,
    );
    assert_eq!(out, vec!["['x.txt']"]);
}

#[test]
fn test_filecmp_dircmp_left_list() {
    let out = run_python(
        r#"
import filecmp, tempfile, os, shutil
d1 = tempfile.mkdtemp()
d2 = tempfile.mkdtemp()
with open(os.path.join(d1, "a.txt"), "w") as f: f.write("a")
with open(os.path.join(d1, "b.txt"), "w") as f: f.write("b")
dc = filecmp.dircmp(d1, d2)
print(sorted(dc.left_list))
shutil.rmtree(d1); shutil.rmtree(d2)
"#,
    );
    assert_eq!(out, vec!["['a.txt', 'b.txt']"]);
}

#[test]
fn test_filecmp_dircmp_right_list() {
    let out = run_python(
        r#"
import filecmp, tempfile, os, shutil
d1 = tempfile.mkdtemp()
d2 = tempfile.mkdtemp()
with open(os.path.join(d2, "z.txt"), "w") as f: f.write("z")
dc = filecmp.dircmp(d1, d2)
print(dc.right_list)
shutil.rmtree(d1); shutil.rmtree(d2)
"#,
    );
    assert_eq!(out, vec!["['z.txt']"]);
}

#[test]
fn test_filecmp_dircmp_left_only() {
    let out = run_python(
        r#"
import filecmp, tempfile, os, shutil
d1 = tempfile.mkdtemp()
d2 = tempfile.mkdtemp()
with open(os.path.join(d1, "only_left.txt"), "w") as f: f.write("x")
dc = filecmp.dircmp(d1, d2)
print(dc.left_only)
shutil.rmtree(d1); shutil.rmtree(d2)
"#,
    );
    assert_eq!(out, vec!["['only_left.txt']"]);
}

#[test]
fn test_filecmp_dircmp_right_only() {
    let out = run_python(
        r#"
import filecmp, tempfile, os, shutil
d1 = tempfile.mkdtemp()
d2 = tempfile.mkdtemp()
with open(os.path.join(d2, "only_right.txt"), "w") as f: f.write("x")
dc = filecmp.dircmp(d1, d2)
print(dc.right_only)
shutil.rmtree(d1); shutil.rmtree(d2)
"#,
    );
    assert_eq!(out, vec!["['only_right.txt']"]);
}

#[test]
fn test_filecmp_dircmp_common() {
    let out = run_python(
        r#"
import filecmp, tempfile, os, shutil
d1 = tempfile.mkdtemp()
d2 = tempfile.mkdtemp()
for d in [d1, d2]:
    with open(os.path.join(d, "shared.txt"), "w") as f: f.write("data")
dc = filecmp.dircmp(d1, d2)
print(dc.common)
shutil.rmtree(d1); shutil.rmtree(d2)
"#,
    );
    assert_eq!(out, vec!["['shared.txt']"]);
}

#[test]
fn test_filecmp_dircmp_same_files() {
    let out = run_python(
        r#"
import filecmp, tempfile, os, shutil
d1 = tempfile.mkdtemp()
d2 = tempfile.mkdtemp()
for d in [d1, d2]:
    with open(os.path.join(d, "same.txt"), "w") as f: f.write("identical")
dc = filecmp.dircmp(d1, d2)
print(dc.same_files)
shutil.rmtree(d1); shutil.rmtree(d2)
"#,
    );
    assert_eq!(out, vec!["['same.txt']"]);
}

#[test]
fn test_filecmp_dircmp_diff_files() {
    let out = run_python(
        r#"
import filecmp, tempfile, os, shutil
d1 = tempfile.mkdtemp()
d2 = tempfile.mkdtemp()
with open(os.path.join(d1, "f.txt"), "w") as f: f.write("aaa")
with open(os.path.join(d2, "f.txt"), "w") as f: f.write("bbb")
dc = filecmp.dircmp(d1, d2)
print(dc.diff_files)
shutil.rmtree(d1); shutil.rmtree(d2)
"#,
    );
    assert_eq!(out, vec!["['f.txt']"]);
}

#[test]
fn test_filecmp_dircmp_subdirs() {
    let out = run_python(
        r#"
import filecmp, tempfile, os, shutil
d1 = tempfile.mkdtemp()
d2 = tempfile.mkdtemp()
for d in [d1, d2]:
    os.makedirs(os.path.join(d, "sub"))
dc = filecmp.dircmp(d1, d2)
print("sub" in dc.subdirs)
shutil.rmtree(d1); shutil.rmtree(d2)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_filecmp_clear_cache_no_error() {
    let out = run_python(
        r#"
import filecmp
filecmp.clear_cache()
print("ok")
"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn test_filecmp_cmp_empty_files_equal() {
    let out = run_python(
        r#"
import filecmp, tempfile, os
d = tempfile.mkdtemp()
a = os.path.join(d, "a.txt")
b = os.path.join(d, "b.txt")
open(a, "w").close()
open(b, "w").close()
print(filecmp.cmp(a, b, shallow=False))
import shutil; shutil.rmtree(d)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_filecmp_cmp_binary_content_matches() {
    let out = run_python(
        r#"
import filecmp, tempfile, os, shutil
d = tempfile.mkdtemp()
a = os.path.join(d, "a.bin")
b = os.path.join(d, "b.bin")
for p in [a, b]:
    with open(p, "wb") as f:
        f.write(bytes(range(256)))
print(filecmp.cmp(a, b, shallow=False))
shutil.rmtree(d)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_filecmp_cmp_binary_content_differs() {
    let out = run_python(
        r#"
import filecmp, tempfile, os, shutil
d = tempfile.mkdtemp()
a = os.path.join(d, "a.bin")
b = os.path.join(d, "b.bin")
with open(a, "wb") as f: f.write(b"\x00\x01\x02")
with open(b, "wb") as f: f.write(b"\x00\x01\x03")
print(filecmp.cmp(a, b, shallow=False))
shutil.rmtree(d)
"#,
    );
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_filecmp_dircmp_empty_dirs_equal() {
    let out = run_python(
        r#"
import filecmp, tempfile, shutil
d1 = tempfile.mkdtemp()
d2 = tempfile.mkdtemp()
dc = filecmp.dircmp(d1, d2)
print(dc.left_only)
print(dc.right_only)
print(dc.diff_files)
shutil.rmtree(d1); shutil.rmtree(d2)
"#,
    );
    assert_eq!(out, vec!["[]", "[]", "[]"]);
}

#[test]
fn test_filecmp_dircmp_common_files_vs_dirs() {
    let out = run_python(
        r#"
import filecmp, tempfile, os, shutil
d1 = tempfile.mkdtemp()
d2 = tempfile.mkdtemp()
# file in d1, dir with same name in d2
with open(os.path.join(d1, "x"), "w") as f: f.write("file")
os.makedirs(os.path.join(d2, "x"))
dc = filecmp.dircmp(d1, d2)
print("x" in dc.common_funny)
shutil.rmtree(d1); shutil.rmtree(d2)
"#,
    );
    assert_eq!(out, vec!["True"]);
}
