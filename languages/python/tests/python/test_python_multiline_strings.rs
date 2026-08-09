// Python multiline strings — triple-quoted, dedent, textwrap, escape sequences
use super::helpers::run_python;

#[test]
fn test_triple_quoted_string() {
    let script = r#"
s = """line one
line two
line three"""
lines = s.split('\n')
print(len(lines))
print(lines[0])
"#;
    assert_eq!(run_python(script), vec!["3", "line one"]);
}

#[test]
fn test_triple_quoted_preserves_whitespace() {
    let script = r#"
s = """
  indented
  content
"""
print(s.strip())
"#;
    assert_eq!(run_python(script), vec!["indented\n  content"]);
}

#[test]
fn test_raw_string_no_escape() {
    let script = r#"
s = r"\n\t\r"
print(len(s))
print(s[0])
"#;
    assert_eq!(run_python(script), vec!["6", "\\"]);
}

#[test]
fn test_raw_triple_quoted() {
    let script = r##"
s = r"""line\none"""
print(s)
"##;
    assert_eq!(run_python(script), vec!["line\\none"]);
}

#[test]
fn test_escape_sequences() {
    let script = r#"
print("tab:\there")
print("newline:\nhere")
print("backslash:\\")
print("quote:\"")
"#;
    assert_eq!(
        run_python(script),
        vec!["tab:\there", "newline:", "here", "backslash:\\", "quote:\""]
    );
}

#[test]
fn test_string_concatenation_literals() {
    let script = r#"
s = ("hello "
     "world "
     "python")
print(s)
"#;
    assert_eq!(run_python(script), vec!["hello world python"]);
}

#[test]
fn test_textwrap_dedent() {
    let script = r#"
import textwrap
s = """\
    line one
    line two
    """
result = textwrap.dedent(s).strip()
print(result)
"#;
    assert_eq!(run_python(script), vec!["line one\nline two"]);
}
