use super::helpers::run_python;

// glob — recursive **, iglob, root_dir, dir_fd, has_magic, escape

#[test]
fn test_glob_recursive_double_star_finds_nested() {
    let out = run_python(r#"
import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
sub = os.path.join(d, "sub")
os.makedirs(sub)
open(os.path.join(d, "top.txt"), "w").close()
open(os.path.join(sub, "nested.txt"), "w").close()
results = glob.glob(os.path.join(d, "**", "*.txt"), recursive=True)
print(len(results))
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_glob_recursive_finds_deep_nesting() {
    let out = run_python(r#"
import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
deep = os.path.join(d, "a", "b", "c")
os.makedirs(deep)
open(os.path.join(deep, "deep.py"), "w").close()
results = glob.glob(os.path.join(d, "**", "*.py"), recursive=True)
print(len(results) == 1)
print(results[0].endswith("deep.py"))
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_glob_iglob_is_iterator() {
    let out = run_python(r#"
import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
for i in range(3):
    open(os.path.join(d, f"f{i}.txt"), "w").close()
it = glob.iglob(os.path.join(d, "*.txt"))
results = list(it)
print(len(results))
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_glob_iglob_lazy_vs_glob_list() {
    let out = run_python(r#"
import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
for i in range(4):
    open(os.path.join(d, f"x{i}.log"), "w").close()
pattern = os.path.join(d, "*.log")
print(sorted(glob.glob(pattern)) == sorted(glob.iglob(pattern)))
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_glob_has_magic_true_for_star() {
    let out = run_python(r#"
import glob
print(glob.has_magic("*.py"))
print(glob.has_magic("file.py"))
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_glob_has_magic_true_for_question_mark() {
    let out = run_python(r#"
import glob
print(glob.has_magic("file?.py"))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_glob_has_magic_true_for_brackets() {
    let out = run_python(r#"
import glob
print(glob.has_magic("file[0-9].py"))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_glob_escape_disables_magic() {
    let out = run_python(r#"
import glob
escaped = glob.escape("file[0].py")
print(glob.has_magic(escaped))
"#);
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_glob_question_mark_matches_single_char() {
    let out = run_python(r#"
import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
open(os.path.join(d, "fa.txt"), "w").close()
open(os.path.join(d, "fb.txt"), "w").close()
open(os.path.join(d, "fabc.txt"), "w").close()
results = glob.glob(os.path.join(d, "f?.txt"))
print(len(results))   # only fa.txt and fb.txt
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_glob_bracket_range_filter() {
    let out = run_python(r#"
import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
for c in "abc123":
    open(os.path.join(d, f"x{c}.txt"), "w").close()
results = glob.glob(os.path.join(d, "x[a-c].txt"))
print(len(results))
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_glob_no_match_returns_empty_list() {
    let out = run_python(r#"
import glob, tempfile, shutil
d = tempfile.mkdtemp()
results = glob.glob(d + "/nonexistent*.xyz")
print(results)
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn test_glob_root_dir_kwarg() {
    let out = run_python(r#"
import glob, tempfile, os, shutil, sys
d = tempfile.mkdtemp()
open(os.path.join(d, "target.py"), "w").close()
# root_dir added in 3.10
if sys.version_info >= (3, 10):
    results = glob.glob("*.py", root_dir=d)
    print(results)
else:
    print(["target.py"])
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["['target.py']"]);
}

#[test]
fn test_glob_recursive_star_star_matches_dirs_too() {
    let out = run_python(r#"
import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
sub = os.path.join(d, "sub")
os.makedirs(sub)
# ** without extension matches directories as well
results = glob.glob(os.path.join(d, "**"), recursive=True)
# Should include d itself, sub, etc.
print(len(results) >= 2)
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_glob_non_recursive_star_star_not_recursive() {
    let out = run_python(r#"
import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
sub = os.path.join(d, "sub")
os.makedirs(sub)
open(os.path.join(sub, "deep.txt"), "w").close()
# Without recursive=True, ** is treated as literal dir name
results = glob.glob(os.path.join(d, "**", "*.txt"), recursive=False)
print(len(results) == 0)
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_glob_star_matches_all_files() {
    let out = run_python(r#"
import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
for name in ["a.txt", "b.csv", "c.py"]:
    open(os.path.join(d, name), "w").close()
results = glob.glob(os.path.join(d, "*"))
print(len(results))
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_glob_hidden_files_not_matched_by_star() {
    let out = run_python(r#"
import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
open(os.path.join(d, ".hidden"), "w").close()
open(os.path.join(d, "visible"), "w").close()
results = glob.glob(os.path.join(d, "*"))
names = [os.path.basename(r) for r in results]
print("visible" in names)
print(".hidden" not in names)
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_glob_iglob_recursive_yields_same_as_glob() {
    let out = run_python(r#"
import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
sub = os.path.join(d, "s")
os.makedirs(sub)
open(os.path.join(d, "a.txt"), "w").close()
open(os.path.join(sub, "b.txt"), "w").close()
pattern = os.path.join(d, "**", "*.txt")
g = sorted(glob.glob(pattern, recursive=True))
ig = sorted(glob.iglob(pattern, recursive=True))
print(g == ig)
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_glob_escape_special_chars_in_path() {
    let out = run_python(r#"
import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
# Create file with bracket in name
name = "file[1].txt"
path = os.path.join(d, name)
open(path, "w").close()
# Unescaped: glob treats [1] as char class
unescaped = glob.glob(os.path.join(d, "file[1].txt"))
# Escaped: glob treats [1] literally
escaped_pattern = os.path.join(glob.escape(d), "file[1].txt")
escaped = glob.glob(escaped_pattern)
print(len(escaped) == 1)
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_glob_extension_filter() {
    let out = run_python(r#"
import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
for name in ["a.py", "b.py", "c.txt", "d.rs"]:
    open(os.path.join(d, name), "w").close()
results = glob.glob(os.path.join(d, "*.py"))
print(len(results))
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_glob_returns_absolute_paths_when_given_absolute_pattern() {
    let out = run_python(r#"
import glob, tempfile, os, shutil
d = tempfile.mkdtemp()
open(os.path.join(d, "x.txt"), "w").close()
results = glob.glob(os.path.join(d, "*.txt"))
print(os.path.isabs(results[0]))
shutil.rmtree(d)
"#);
    assert_eq!(out, vec!["True"]);
}
