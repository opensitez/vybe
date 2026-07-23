use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: re (regex) — search, match, findall, groups, flags, sub, split, lookaheads, compile
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_re_search_and_match() {
    let src = r#"
import re

print(re.search(r"\d+", "abc 123 def").group())
print(re.match(r"\d+", "123abc").group())
print(re.match(r"\d+", "abc123") is None)
print(re.fullmatch(r"\d{3}", "123") is not None)
print(re.fullmatch(r"\d{3}", "1234") is None)
"#;
    assert_eq!(run_python(src), vec!["123", "123", "True", "True", "True"]);
}

#[test]
fn test_py_re_findall_and_finditer() {
    let src = r#"
import re

text = "The price is $10 and $25.50"
amounts = re.findall(r'\$(\d+(?:\.\d+)?)', text)
print(amounts)

positions = [(m.start(), m.group()) for m in re.finditer(r'\$\d+', text)]
print(positions)
"#;
    assert_eq!(
        run_python(src),
        vec!["['10', '25.50']", "[(13, '$10'), (21, '$25')]"]
    );
}

#[test]
fn test_py_re_groups_and_named_groups() {
    let src = r#"
import re

m = re.search(r'(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})', "Event: 2024-03-15")
print(m.group('year'))
print(m.group('month'))
print(m.group('day'))
print(m.groupdict())
"#;
    assert_eq!(
        run_python(src),
        vec![
            "2024",
            "03",
            "15",
            "{'year': '2024', 'month': '03', 'day': '15'}"
        ]
    );
}

#[test]
fn test_py_re_sub_replacement() {
    let src = r#"
import re

result = re.sub(r'\d+', 'NUM', "I have 3 cats and 12 dogs")
print(result)

result2 = re.sub(r'\d+', lambda m: str(int(m.group()) * 2), "Score: 10 + 5")
print(result2)
"#;
    assert_eq!(
        run_python(src),
        vec!["I have NUM cats and NUM dogs", "Score: 20 + 10"]
    );
}

#[test]
fn test_py_re_split_pattern() {
    let src = r#"
import re

parts = re.split(r'[\s,;]+', "hello world,foo;bar  baz")
print(parts)

parts2 = re.split(r'(\s+)', "one two  three")
print(parts2)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "['hello', 'world', 'foo', 'bar', 'baz']",
            "['one', ' ', 'two', '  ', 'three']"
        ]
    );
}

#[test]
fn test_py_re_flags_ignorecase_multiline() {
    let src = r#"
import re

print(re.search(r'hello', 'HELLO WORLD', re.IGNORECASE).group())
text = "line1\nline2\nline3"
starts = re.findall(r'^line\d', text, re.MULTILINE)
print(starts)
"#;
    assert_eq!(
        run_python(src),
        vec!["HELLO", "['line1', 'line2', 'line3']"]
    );
}

#[test]
fn test_py_re_dotall_and_verbose_flag() {
    let src = r#"
import re

m = re.search(r'start.+end', "start\nmiddle\nend", re.DOTALL)
print(m is not None)

email_pattern = re.compile(r'''
    [\w.+-]+    # username
    @           # at sign
    [\w-]+      # domain
    \.[a-z]+    # tld
''', re.VERBOSE)
print(email_pattern.search("Contact: user@example.com here").group())
"#;
    assert_eq!(run_python(src), vec!["True", "user@example.com"]);
}

#[test]
fn test_py_re_lookahead_lookbehind() {
    let src = r#"
import re

# Positive lookahead
prices = re.findall(r'\d+(?= USD)', "10 USD and 20 EUR and 30 USD")
print(prices)

# Negative lookahead
non_usd = re.findall(r'\d+(?! USD)', "10 USD and 20 EUR and 30 USD")
print(non_usd)

# Positive lookbehind
usd_values = re.findall(r'(?<=\$)\d+', "$100 and $200")
print(usd_values)
"#;
    assert_eq!(
        run_python(src),
        vec!["['10', '30']", "['20', '30']", "['100', '200']"]
    );
}

#[test]
fn test_py_re_compile_and_reuse() {
    let src = r#"
import re

pattern = re.compile(r'\b\w{5}\b')
texts = ["Hello world", "The quick brown fox", "abc defgh"]
for t in texts:
    found = pattern.findall(t)
    print(found if found else [])
"#;
    assert_eq!(
        run_python(src),
        vec!["['Hello', 'world']", "['quick', 'brown']", "['defgh']"]
    );
}

#[test]
fn test_py_re_backreferences() {
    let src = r#"
import re

# Match repeated words
m = re.search(r'\b(\w+)\s+\1\b', "the the fox jumps")
print(m.group() if m else None)

# XML-like matched tags
m2 = re.search(r'<(\w+)>.*?</\1>', "<div>content</div>")
print(m2.group())
"#;
    assert_eq!(run_python(src), vec!["the the", "<div>content</div>"]);
}

#[test]
fn test_py_re_non_capturing_group() {
    let src = r#"
import re

# Non-capturing group (?:...)
m = re.search(r'(?:foo|bar) (\w+)', "foo baz")
print(m.group(1))  # only one captured group

m2 = re.search(r'(?:foo|bar) (\w+)', "bar qux")
print(m2.group(1))
"#;
    assert_eq!(run_python(src), vec!["baz", "qux"]);
}

#[test]
fn test_py_re_span_start_end() {
    let src = r#"
import re

m = re.search(r'world', "hello world!")
print(m.start())
print(m.end())
print(m.span())
"#;
    assert_eq!(run_python(src), vec!["6", "11", "(6, 11)"]);
}

#[test]
fn test_py_re_escape_special_characters() {
    let src = r#"
import re

user_input = "2+3=5 (maybe?)"
escaped = re.escape(user_input)
print(re.search(escaped, user_input) is not None)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_re_subn_count() {
    let src = r#"
import re

result, count = re.subn(r'\d', 'X', "abc123def456")
print(result)
print(count)
"#;
    assert_eq!(run_python(src), vec!["abcXXXdefXXX", "6"]);
}

#[test]
fn test_py_re_pattern_unicode_and_word_boundary() {
    let src = r#"
import re

words = re.findall(r'\b\w+\b', "Hello, world! How are you?")
print(words)

# Unicode word matching
print(re.search(r'\w+', 'café').group())
"#;
    assert_eq!(
        run_python(src),
        vec!["['Hello', 'world', 'How', 'are', 'you']", "café"]
    );
}
