use super::helpers::run_python;

// linecache — getline, getlines, checkcache, clearcache, updatecache, lazycache

#[test]
fn test_linecache_getline_existing_line() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.write("line one\nline two\nline three\n")
f.close()
linecache.clearcache()
result = linecache.getline(f.name, 2)
print(result.strip())
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["line two"]);
}

#[test]
fn test_linecache_getline_out_of_range_returns_empty() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.write("only one line\n")
f.close()
linecache.clearcache()
result = linecache.getline(f.name, 99)
print(repr(result))
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["''"]);
}

#[test]
fn test_linecache_getlines_returns_all_lines() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.write("a\nb\nc\n")
f.close()
linecache.clearcache()
lines = linecache.getlines(f.name)
print(lines)
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["['a\\n', 'b\\n', 'c\\n']"]);
}

#[test]
fn test_linecache_getlines_nonexistent_returns_empty() {
    let out = run_python(
        r#"
import linecache
lines = linecache.getlines("/nonexistent/path/xyz.py")
print(lines)
"#,
    );
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn test_linecache_clearcache_empties_cache() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.write("cached line\n")
f.close()
linecache.getline(f.name, 1)   # populate cache
linecache.clearcache()
print(len(linecache.cache) == 0)
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_linecache_cache_populated_after_getline() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.write("something\n")
f.close()
linecache.clearcache()
linecache.getline(f.name, 1)
print(f.name in linecache.cache)
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_linecache_cache_entry_structure() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.write("line1\nline2\n")
f.close()
linecache.clearcache()
linecache.getline(f.name, 1)
entry = linecache.cache[f.name]
# cache entry: (size, mtime, lines, fullname)
print(len(entry) == 4)
print(isinstance(entry[2], list))
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_linecache_checkcache_after_file_changed() {
    let out = run_python(
        r#"
import linecache, tempfile, os, time
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.write("original\n")
f.close()
linecache.clearcache()
linecache.getline(f.name, 1)
time.sleep(0.01)
with open(f.name, "w") as fh:
    fh.write("changed\n")
linecache.checkcache(f.name)
result = linecache.getline(f.name, 1)
print(result.strip())
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["changed"]);
}

#[test]
fn test_linecache_getline_line_1_is_first() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.write("first\nsecond\n")
f.close()
linecache.clearcache()
print(linecache.getline(f.name, 1).strip())
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["first"]);
}

#[test]
fn test_linecache_getline_last_line() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.write("alpha\nbeta\ngamma\n")
f.close()
linecache.clearcache()
print(linecache.getline(f.name, 3).strip())
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["gamma"]);
}

#[test]
fn test_linecache_getlines_preserves_newlines() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.write("x\ny\n")
f.close()
linecache.clearcache()
lines = linecache.getlines(f.name)
print(all(l.endswith("\n") for l in lines))
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_linecache_getline_zero_returns_empty() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.write("data\n")
f.close()
linecache.clearcache()
print(repr(linecache.getline(f.name, 0)))
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["''"]);
}

#[test]
fn test_linecache_updatecache_explicit() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.write("hello\n")
f.close()
linecache.clearcache()
linecache.updatecache(f.name)
print(f.name in linecache.cache)
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_linecache_lazycache_for_module() {
    let out = run_python(
        r#"
import linecache, sys
# lazycache works for modules with __loader__
modname = "os"
linecache.clearcache()
result = linecache.lazycache(modname, vars(sys.modules[modname]))
# Returns True if a lazy entry was created, False if already cached
print(isinstance(result, bool))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_linecache_getlines_count_matches_file() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
for i in range(10):
    f.write(f"line {i}\n")
f.close()
linecache.clearcache()
lines = linecache.getlines(f.name)
print(len(lines))
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_linecache_getline_includes_trailing_newline() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.write("test line\n")
f.close()
linecache.clearcache()
line = linecache.getline(f.name, 1)
print(line.endswith("\n"))
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_linecache_negative_lineno_returns_empty() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.write("data\n")
f.close()
linecache.clearcache()
print(repr(linecache.getline(f.name, -1)))
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["''"]);
}

#[test]
fn test_linecache_getlines_empty_file() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.close()
linecache.clearcache()
print(linecache.getlines(f.name))
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn test_linecache_checkcache_with_no_args_clears_stale() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
f.write("data\n")
f.close()
linecache.clearcache()
linecache.getline(f.name, 1)
os.unlink(f.name)
# File deleted — checkcache() should remove stale entry
linecache.checkcache()
print(f.name not in linecache.cache)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_linecache_getline_multiline_file() {
    let out = run_python(
        r#"
import linecache, tempfile, os
f = tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".py")
lines = [f"row{i}\n" for i in range(5)]
f.write("".join(lines))
f.close()
linecache.clearcache()
for i, expected in enumerate(lines, start=1):
    got = linecache.getline(f.name, i)
    assert got == expected, f"line {i}: {got!r} != {expected!r}"
print("all ok")
os.unlink(f.name)
"#,
    );
    assert_eq!(out, vec!["all ok"]);
}
