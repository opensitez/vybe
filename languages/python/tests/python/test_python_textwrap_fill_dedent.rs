use super::helpers::run_python;

// textwrap — wrap, fill, dedent, indent, shorten, TextWrapper

#[test]
fn test_textwrap_wrap_returns_list_of_lines() {
    let out = run_python(
        r#"
import textwrap
lines = textwrap.wrap("one two three four five six", width=15)
print(lines)
"#,
    );
    assert_eq!(out, vec!["['one two three', 'four five six']"]);
}

#[test]
fn test_textwrap_fill_joins_with_newline() {
    let out = run_python(
        r#"
import textwrap
result = textwrap.fill("one two three four five six", width=15)
print(result)
"#,
    );
    assert_eq!(out, vec!["one two three", "four five six"]);
}

#[test]
fn test_textwrap_dedent_removes_common_leading_whitespace() {
    let out = run_python(
        r#"
import textwrap
text = "    line1\n    line2\n    line3"
print(textwrap.dedent(text))
"#,
    );
    assert_eq!(out, vec!["line1", "line2", "line3"]);
}

#[test]
fn test_textwrap_dedent_partial_indent() {
    let out = run_python(
        r#"
import textwrap
text = "    line1\n  line2"
print(repr(textwrap.dedent(text)))
"#,
    );
    assert_eq!(out, vec!["'  line1\\nline2'"]);
}

#[test]
fn test_textwrap_indent_adds_prefix_to_all_lines() {
    let out = run_python(
        r#"
import textwrap
result = textwrap.indent("line1\nline2", prefix="  ")
print(result)
"#,
    );
    assert_eq!(out, vec!["  line1", "  line2"]);
}

#[test]
fn test_textwrap_indent_with_predicate_skips_empty() {
    let out = run_python(
        r#"
import textwrap
result = textwrap.indent("line1\n\nline2", prefix="> ", predicate=lambda s: s.strip())
print(result)
"#,
    );
    assert_eq!(out, vec!["> line1", "", "> line2"]);
}

#[test]
fn test_textwrap_shorten_truncates_to_width() {
    let out = run_python(
        r#"
import textwrap
result = textwrap.shorten("Hello world this is a long sentence", width=20)
print(result)
"#,
    );
    assert_eq!(out, vec!["Hello world [...]"]);
}

#[test]
fn test_textwrap_shorten_custom_placeholder() {
    let out = run_python(
        r#"
import textwrap
result = textwrap.shorten("Hello world this is long", width=15, placeholder="...")
print(result)
"#,
    );
    assert_eq!(out, vec!["Hello world..."]);
}

#[test]
fn test_textwrap_textwrapper_initial_indent() {
    let out = run_python(
        r#"
import textwrap
tw = textwrap.TextWrapper(width=20, initial_indent=">>> ", subsequent_indent="    ")
lines = tw.wrap("one two three four five six seven")
print(lines[0])
print(lines[1].startswith("    "))
"#,
    );
    assert_eq!(out, vec![">>> one two three", "True"]);
}

#[test]
fn test_textwrap_textwrapper_break_long_words_false() {
    let out = run_python(
        r#"
import textwrap
tw = textwrap.TextWrapper(width=5, break_long_words=False)
result = tw.wrap("short averylongword")
print(result)
"#,
    );
    assert_eq!(out, vec!["['short', 'averylongword']"]);
}

#[test]
fn test_textwrap_textwrapper_break_long_words_true() {
    let out = run_python(
        r#"
import textwrap
tw = textwrap.TextWrapper(width=5, break_long_words=True)
result = tw.wrap("abcdefghij")
print(result)
"#,
    );
    assert_eq!(out, vec!["['abcde', 'fghij']"]);
}

#[test]
fn test_textwrap_wrap_empty_string() {
    let out = run_python(
        r#"
import textwrap
print(textwrap.wrap(""))
"#,
    );
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn test_textwrap_wrap_single_word_fits() {
    let out = run_python(
        r#"
import textwrap
print(textwrap.wrap("hello", width=10))
"#,
    );
    assert_eq!(out, vec!["['hello']"]);
}

#[test]
fn test_textwrap_fill_preserves_paragraph_structure() {
    let out = run_python(
        r#"
import textwrap
text = "aaa bbb ccc ddd eee"
result = textwrap.fill(text, width=10)
for line in result.splitlines():
    print(len(line) <= 10)
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_textwrap_dedent_no_common_indent() {
    let out = run_python(
        r#"
import textwrap
text = "no\n  indent\nhere"
print(textwrap.dedent(text))
"#,
    );
    assert_eq!(out, vec!["no", "  indent", "here"]);
}

#[test]
fn test_textwrap_textwrapper_max_lines() {
    let out = run_python(
        r#"
import textwrap
tw = textwrap.TextWrapper(width=10, max_lines=2, placeholder=" ...")
lines = tw.wrap("one two three four five six seven eight nine")
print(len(lines))
print(lines[-1].endswith("..."))
"#,
    );
    assert_eq!(out, vec!["2", "True"]);
}

#[test]
fn test_textwrap_wrap_break_on_hyphens_true() {
    let out = run_python(
        r#"
import textwrap
tw = textwrap.TextWrapper(width=8, break_on_hyphens=True)
result = tw.wrap("pre-fix")
print(len(result) >= 1)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_textwrap_wrap_expand_tabs() {
    let out = run_python(
        r#"
import textwrap
tw = textwrap.TextWrapper(width=40, expand_tabs=True, tabsize=4)
result = tw.fill("\thello world")
print(result.startswith("    hello"))
"#,
    );
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_textwrap_fill_drop_whitespace_false() {
    let out = run_python(
        r#"
import textwrap
result = textwrap.fill("a  b  c", width=40, drop_whitespace=False)
print(result)
"#,
    );
    assert_eq!(out, vec!["a  b  c"]);
}

#[test]
fn test_textwrap_indent_empty_string() {
    let out = run_python(
        r#"
import textwrap
print(repr(textwrap.indent("", prefix="  ")))
"#,
    );
    assert_eq!(out, vec!["''"]);
}
