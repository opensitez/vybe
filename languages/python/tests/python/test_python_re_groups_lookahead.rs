// Python re groups, named groups, lookahead, lookbehind, non-capturing
use super::helpers::run_python;

#[test]
fn test_re_named_groups() {
    let script = r#"
import re
m = re.match(r'(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})', '2024-07-15')
print(m.group('year'))
print(m.group('month'))
print(m.group('day'))
"#;
    assert_eq!(run_python(script), vec!["2024", "07", "15"]);
}

#[test]
fn test_re_positive_lookahead() {
    let script = r#"
import re
result = re.findall(r'\d+(?= dollars)', '100 dollars and 200 euros')
print(result)
"#;
    assert_eq!(run_python(script), vec!["['100']"]);
}

#[test]
fn test_re_negative_lookahead() {
    let script = r#"
import re
result = re.findall(r'\d+(?! dollars)', '100 dollars 200 euros')
print(result)
"#;
    assert_eq!(run_python(script), vec!["['200']"]);
}

#[test]
fn test_re_lookbehind() {
    let script = r#"
import re
result = re.findall(r'(?<=\$)\d+', '$100 $200 €300')
print(result)
"#;
    assert_eq!(run_python(script), vec!["['100', '200']"]);
}

#[test]
fn test_re_non_capturing_group() {
    let script = r#"
import re
m = re.match(r'(?:foo|bar)(baz)', 'foobaz')
print(m.group(1))
print(m.lastindex)
"#;
    assert_eq!(run_python(script), vec!["baz", "1"]);
}

#[test]
fn test_re_backreference() {
    let script = r#"
import re
result = re.findall(r'(\w+) \1', 'hello hello world world test')
print(result)
"#;
    assert_eq!(run_python(script), vec!["['hello', 'world']"]);
}

#[test]
fn test_re_sub_with_group() {
    let script = r#"
import re
result = re.sub(r'(\w+) (\w+)', r'\2 \1', 'hello world')
print(result)
"#;
    assert_eq!(run_python(script), vec!["world hello"]);
}

#[test]
fn test_re_finditer_spans() {
    let script = r#"
import re
spans = [(m.start(), m.end()) for m in re.finditer(r'\d+', 'abc123def456')]
print(spans)
"#;
    assert_eq!(run_python(script), vec!["[(3, 6), (9, 12)]"]);
}

#[test]
fn test_re_flags_ignorecase() {
    let script = r#"
import re
m = re.search(r'hello', 'HELLO WORLD', re.IGNORECASE)
print(m is not None)
print(m.group())
"#;
    assert_eq!(run_python(script), vec!["True", "HELLO"]);
}
