use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: String Regular Expressions Advanced — re.compile, search, match, findall, sub, named groups, lookaround
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_re_search_named_capturing_groups() {
    let src = r#"
import re

pattern = r"(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})"
match = re.search(pattern, "Date: 2024-05-12")

print(match.group("year"))
print(match.group("month"))
print(match.group("day"))
print(match.groupdict())
"#;
    assert_eq!(
        run_python(src),
        vec![
            "2024",
            "05",
            "12",
            "{'year': '2024', 'month': '05', 'day': '12'}"
        ]
    );
}

#[test]
fn test_py_re_sub_with_replacement_callable() {
    let src = r#"
import re

def uppercase_match(m):
    return m.group(0).upper()

text = "hello world python"
result = re.sub(r"\b\w+\b", uppercase_match, text)
print(result)
"#;
    assert_eq!(run_python(src), vec!["HELLO WORLD PYTHON"]);
}

#[test]
fn test_py_re_finditer_match_span_tuples() {
    let src = r#"
import re

text = "cat, bat, rat, mat"
matches = list(re.finditer(r"\b\w+at\b", text))
print([m.group() for m in matches])
print([m.span() for m in matches])
"#;
    assert_eq!(
        run_python(src),
        vec![
            "['cat', 'bat', 'rat', 'mat']",
            "[(0, 3), (5, 8), (10, 13), (15, 18)]"
        ]
    );
}

#[test]
fn test_py_re_lookahead_lookbehind_assertions() {
    let src = r#"
import re

# Positive lookahead: digits followed by 'px'
px_vals = re.findall(r"\d+(?=px)", "width: 100px, height: 200px, id: 300")
print(px_vals)

# Positive lookbehind: digits preceded by '$'
prices = re.findall(r"(?<=\$)\d+", "Items: $10, $25, #50")
print(prices)
"#;
    assert_eq!(run_python(src), vec!["['100', '200']", "['10', '25']"]);
}

#[test]
fn test_py_re_flags_ignorecase_multiline_dotall() {
    let src = r#"
import re

print(bool(re.search(r"hello", "HELLO", re.IGNORECASE)))

multiline_matches = re.findall(r"^start.*$", "start line 1\nmiddle line\nstart line 2", re.MULTILINE)
print(multiline_matches)

dotall_match = re.search(r"a.*b", "a\n\nb", re.DOTALL)
print(dotall_match.group())
"#;
    assert_eq!(
        run_python(src),
        vec!["True", "['start line 1', 'start line 2']", "a\n\nb"]
    );
}

#[test]
fn test_py_re_split_with_capturing_parentheses() {
    let src = r#"
import re

parts = re.split(r"([,;])", "apple,banana;cherry")
print(parts)
"#;
    assert_eq!(
        run_python(src),
        vec!["['apple', ',', 'banana', ';', 'cherry']"]
    );
}

#[test]
fn test_py_re_escape_literal_special_characters() {
    let src = r#"
import re

raw_input = "price is $10.00 (tax incl.)"
escaped = re.escape(raw_input)
print(bool(re.search(escaped, raw_input)))
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_re_non_capturing_groups() {
    let src = r#"
import re

pattern = r"(?:http|https)://(?P<domain>[\w.]+)"
match = re.search(pattern, "https://example.com")
print(match.group("domain"))
print(match.groups())  # non-capturing group not in .groups()!
"#;
    assert_eq!(run_python(src), vec!["example.com", "('example.com',)"]);
}

#[test]
fn test_py_re_subn_replacement_count_tuple() {
    let src = r#"
import re

text = "foo bar foo baz foo"
result, count = re.subn(r"foo", "qux", text)
print(result)
print(count)
"#;
    assert_eq!(run_python(src), vec!["qux bar qux baz qux", "3"]);
}

#[test]
fn test_py_re_compiled_pattern_reuse() {
    let src = r#"
import re

regex = re.compile(r"\b[A-Z]\w+\b")
print(regex.findall("Alice and Bob went to Charlie"))
"#;
    assert_eq!(run_python(src), vec!["['Alice', 'Bob', 'Charlie']"]);
}
