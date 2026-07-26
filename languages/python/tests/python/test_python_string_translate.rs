// Python str.translate() and maketrans — character substitution and deletion
use super::helpers::run_python;

#[test]
fn test_translate_basic() {
    let script = r#"
table = str.maketrans('aeiou', '12345')
s = "hello world"
print(s.translate(table))
"#;
    assert_eq!(run_python(script), vec!["h2ll4 w4rld"]);
}

#[test]
fn test_translate_delete_chars() {
    let script = r#"
table = str.maketrans('', '', 'aeiou')
s = "hello world"
print(s.translate(table))
"#;
    assert_eq!(run_python(script), vec!["hll wrld"]);
}

#[test]
fn test_translate_with_dict() {
    let script = r#"
table = str.maketrans({ord('a'): 'AA', ord('b'): None})
s = "ab ab ab"
print(s.translate(table))
"#;
    assert_eq!(run_python(script), vec!["AA AA AA"]);
}

#[test]
fn test_translate_unicode() {
    let script = r#"
table = str.maketrans('\u00e9\u00e0', 'ea')
s = "caf\u00e9 \u00e0 la mode"
print(s.translate(table))
"#;
    assert_eq!(run_python(script), vec!["cafe a la mode"]);
}

#[test]
fn test_maketrans_three_args() {
    let script = r#"
intab = "aeiou"
outtab = "12345"
deltab = " "
tbl = str.maketrans(intab, outtab, deltab)
s = "hello world"
print(s.translate(tbl))
"#;
    assert_eq!(run_python(script), vec!["h2ll4w4rld"]);
}

#[test]
fn test_translate_no_op() {
    let script = r#"
s = "unchanged"
table = str.maketrans({})
print(s.translate(table))
"#;
    assert_eq!(run_python(script), vec!["unchanged"]);
}
