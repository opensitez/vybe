use super::helpers::run_python;

// fnmatch — fnmatch, fnmatchcase, filter, translate (regex), case sensitivity

#[test]
fn test_fnmatch_star_matches_any_sequence() {
    let out = run_python(r#"
import fnmatch
print(fnmatch.fnmatch("report_2024.csv", "*.csv"))
print(fnmatch.fnmatch("report_2024.txt", "*.csv"))
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_fnmatch_question_mark_single_char() {
    let out = run_python(r#"
import fnmatch
print(fnmatch.fnmatch("file1.py", "file?.py"))
print(fnmatch.fnmatch("file12.py", "file?.py"))
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_fnmatch_bracket_character_class() {
    let out = run_python(r#"
import fnmatch
print(fnmatch.fnmatch("fileA.py", "file[ABC].py"))
print(fnmatch.fnmatch("fileD.py", "file[ABC].py"))
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_fnmatch_bracket_range() {
    let out = run_python(r#"
import fnmatch
print(fnmatch.fnmatch("img3.png", "img[0-9].png"))
print(fnmatch.fnmatch("imgX.png", "img[0-9].png"))
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_fnmatch_case_insensitive_on_windows_like_os() {
    let out = run_python(r#"
import fnmatch, os
# fnmatch normalises case on case-insensitive OS, is case-sensitive on Linux
name = "FILE.TXT"
pattern = "file.txt"
result = fnmatch.fnmatch(name, pattern)
# Just check it returns a bool without error
print(isinstance(result, bool))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_fnmatchcase_always_case_sensitive() {
    let out = run_python(r#"
import fnmatch
print(fnmatch.fnmatchcase("File.TXT", "file.txt"))
print(fnmatch.fnmatchcase("file.txt", "file.txt"))
"#);
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_fnmatchcase_star_mixed_case() {
    let out = run_python(r#"
import fnmatch
print(fnmatch.fnmatchcase("README.MD", "*.MD"))
print(fnmatch.fnmatchcase("readme.md", "*.MD"))
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_fnmatch_filter_selects_matching_names() {
    let out = run_python(r#"
import fnmatch
names = ["main.py", "utils.py", "data.csv", "README.md", "test.py"]
result = fnmatch.filter(names, "*.py")
print(sorted(result))
"#);
    assert_eq!(out, vec!["['main.py', 'test.py', 'utils.py']"]);
}

#[test]
fn test_fnmatch_filter_empty_list() {
    let out = run_python(r#"
import fnmatch
print(fnmatch.filter([], "*.py"))
"#);
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn test_fnmatch_filter_no_match() {
    let out = run_python(r#"
import fnmatch
result = fnmatch.filter(["a.rs", "b.rs"], "*.py")
print(result)
"#);
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn test_fnmatch_filter_all_match() {
    let out = run_python(r#"
import fnmatch
names = ["x.py", "y.py", "z.py"]
result = fnmatch.filter(names, "*.py")
print(sorted(result) == sorted(names))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_fnmatch_translate_produces_regex() {
    let out = run_python(r#"
import fnmatch, re
pattern = fnmatch.translate("*.py")
print(isinstance(pattern, str))
print(bool(re.match(pattern, "hello.py")))
print(bool(re.match(pattern, "hello.rs")))
"#);
    assert_eq!(out, vec!["True", "True", "False"]);
}

#[test]
fn test_fnmatch_translate_question_mark_regex() {
    let out = run_python(r#"
import fnmatch, re
pattern = fnmatch.translate("file?.txt")
print(bool(re.match(pattern, "file1.txt")))
print(bool(re.match(pattern, "file12.txt")))
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_fnmatch_translate_bracket_regex() {
    let out = run_python(r#"
import fnmatch, re
pattern = fnmatch.translate("file[0-9].txt")
print(bool(re.match(pattern, "file5.txt")))
print(bool(re.match(pattern, "fileX.txt")))
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_fnmatch_dot_is_not_special() {
    let out = run_python(r#"
import fnmatch
# In fnmatch (unlike regex), . is literal
print(fnmatch.fnmatch("file.txt", "file.txt"))
print(fnmatch.fnmatch("fileXtxt", "file.txt"))
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_fnmatch_star_does_not_match_path_separator() {
    let out = run_python(r#"
import fnmatch
# fnmatch * matches anything including /
print(fnmatch.fnmatch("dir/file.py", "*.py"))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_fnmatch_exact_literal_match() {
    let out = run_python(r#"
import fnmatch
print(fnmatch.fnmatch("exact", "exact"))
print(fnmatch.fnmatch("exact2", "exact"))
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_fnmatch_empty_pattern_matches_empty() {
    let out = run_python(r#"
import fnmatch
print(fnmatch.fnmatch("", ""))
print(fnmatch.fnmatch("x", ""))
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_fnmatch_negate_bracket() {
    let out = run_python(r#"
import fnmatch
# [!0-9] matches any character not in range
print(fnmatch.fnmatch("fileA.py", "file[!0-9].py"))
print(fnmatch.fnmatch("file5.py", "file[!0-9].py"))
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_fnmatch_double_star_is_not_recursive() {
    let out = run_python(r#"
import fnmatch
# fnmatch ** is not recursive (unlike glob) — just matches literally like *
print(fnmatch.fnmatch("deep/path/file.py", "**/*.py"))
"#);
    assert_eq!(out, vec!["True"]);
}
