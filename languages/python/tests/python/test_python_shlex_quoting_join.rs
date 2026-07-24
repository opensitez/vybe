use super::helpers::run_python;

// shlex — split, quote, join, Shlex class, posix mode, punctuation_chars

#[test]
fn test_shlex_split_basic() {
    let out = run_python(r#"
import shlex
result = shlex.split("one two three")
print(result)
"#);
    assert_eq!(out, vec!["['one', 'two', 'three']"]);
}

#[test]
fn test_shlex_split_quoted_string() {
    let out = run_python(r#"
import shlex
result = shlex.split("hello 'world foo' bar")
print(result)
"#);
    assert_eq!(out, vec!["['hello', 'world foo', 'bar']"]);
}

#[test]
fn test_shlex_split_double_quotes() {
    let out = run_python(r#"
import shlex
result = shlex.split('echo "hello world"')
print(result)
"#);
    assert_eq!(out, vec!["['echo', 'hello world']"]);
}

#[test]
fn test_shlex_split_escaped_space() {
    let out = run_python(r#"
import shlex
result = shlex.split(r"one\ two three")
print(result)
"#);
    assert_eq!(out, vec!["['one two', 'three']"]);
}

#[test]
fn test_shlex_split_posix_false() {
    let out = run_python(r#"
import shlex
result = shlex.split("'hello world'", posix=False)
# posix=False keeps the quotes
print(result)
"#);
    assert_eq!(out, vec!["[\"'hello world'\"]"]);
}

#[test]
fn test_shlex_split_empty_string() {
    let out = run_python(r#"
import shlex
print(shlex.split(""))
"#);
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn test_shlex_quote_safe_string() {
    let out = run_python(r#"
import shlex
print(shlex.quote("hello"))
"#);
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn test_shlex_quote_string_with_spaces() {
    let out = run_python(r#"
import shlex
print(shlex.quote("hello world"))
"#);
    assert_eq!(out, vec!["'hello world'"]);
}

#[test]
fn test_shlex_quote_string_with_special_chars() {
    let out = run_python(r#"
import shlex
result = shlex.quote("rm -rf /")
print(result)
"#);
    assert_eq!(out, vec!["'rm -rf /'"]);
}

#[test]
fn test_shlex_quote_empty_string_becomes_empty_quotes() {
    let out = run_python(r#"
import shlex
print(shlex.quote(""))
"#);
    assert_eq!(out, vec!["''"]);
}

#[test]
fn test_shlex_join_list_to_string() {
    let out = run_python(r#"
import shlex
result = shlex.join(["ls", "-la", "/tmp"])
print(result)
"#);
    assert_eq!(out, vec!["ls -la /tmp"]);
}

#[test]
fn test_shlex_join_quotes_items_with_spaces() {
    let out = run_python(r#"
import shlex
result = shlex.join(["git", "commit", "-m", "fix the bug"])
print(result)
"#);
    assert_eq!(out, vec!["git commit -m 'fix the bug'"]);
}

#[test]
fn test_shlex_split_join_roundtrip() {
    let out = run_python(r#"
import shlex
original = ["echo", "hello world", "foo"]
joined = shlex.join(original)
recovered = shlex.split(joined)
print(recovered == original)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_shlex_class_wordchars() {
    let out = run_python(r#"
import shlex
s = shlex.shlex("hello_world foo", posix=True)
s.whitespace_split = True
tokens = list(s)
print(tokens)
"#);
    assert_eq!(out, vec!["['hello_world', 'foo']"]);
}

#[test]
fn test_shlex_class_token_iteration() {
    let out = run_python(r#"
import shlex
s = shlex.shlex("a b c", posix=True)
tokens = list(s)
print(tokens)
"#);
    assert_eq!(out, vec!["['a', 'b', 'c']"]);
}

#[test]
fn test_shlex_class_comments() {
    let out = run_python(r#"
import shlex, io
src = io.StringIO("hello # this is a comment\nworld")
s = shlex.shlex(src, posix=True)
tokens = list(s)
print(tokens)
"#);
    assert_eq!(out, vec!["['hello', 'world']"]);
}

#[test]
fn test_shlex_class_custom_comment_char() {
    let out = run_python(r#"
import shlex, io
src = io.StringIO("foo ; ignored\nbar")
s = shlex.shlex(src, posix=True)
s.commenters = ";"
tokens = list(s)
print(tokens)
"#);
    assert_eq!(out, vec!["['foo', 'bar']"]);
}

#[test]
fn test_shlex_split_multiple_spaces() {
    let out = run_python(r#"
import shlex
result = shlex.split("a   b    c")
print(result)
"#);
    assert_eq!(out, vec!["['a', 'b', 'c']"]);
}

#[test]
fn test_shlex_split_unclosed_quote_raises() {
    let out = run_python(r#"
import shlex
try:
    shlex.split("'unclosed")
except ValueError:
    print("ValueError")
"#);
    assert_eq!(out, vec!["ValueError"]);
}

#[test]
fn test_shlex_quote_backslash_handling() {
    let out = run_python(r#"
import shlex
result = shlex.quote("file\\name")
print(len(result) > 0)
# After quoting and splitting, we should recover the original
recovered = shlex.split(result)[0]
print(recovered)
"#);
    assert_eq!(out, vec!["True", "file\\name"]);
}
