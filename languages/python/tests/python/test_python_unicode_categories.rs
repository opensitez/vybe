// Python unicode — unicodedata, categories, normalization, str methods
use super::helpers::run_python;

#[test]
fn test_unicode_name() {
    let script = r#"
import unicodedata
print(unicodedata.name('\u00e9'))
print(unicodedata.name('A'))
"#;
    assert_eq!(run_python(script), vec!["LATIN SMALL LETTER E WITH ACUTE", "LATIN CAPITAL LETTER A"]);
}

#[test]
fn test_unicode_category() {
    let script = r#"
import unicodedata
print(unicodedata.category('A'))   # Lu - uppercase letter
print(unicodedata.category('a'))   # Ll - lowercase letter
print(unicodedata.category('1'))   # Nd - decimal digit
print(unicodedata.category(' '))   # Zs - space separator
"#;
    assert_eq!(run_python(script), vec!["Lu", "Ll", "Nd", "Zs"]);
}

#[test]
fn test_unicode_normalize_nfc() {
    let script = r#"
import unicodedata
# composed vs decomposed e-acute
composed = '\u00e9'
decomposed = 'e\u0301'
print(composed == decomposed)
nfc = unicodedata.normalize('NFC', decomposed)
print(nfc == composed)
"#;
    assert_eq!(run_python(script), vec!["False", "True"]);
}

#[test]
fn test_unicode_normalize_nfd() {
    let script = r#"
import unicodedata
s = '\u00e9'
nfd = unicodedata.normalize('NFD', s)
print(len(nfd))
print(nfd[0])
"#;
    assert_eq!(run_python(script), vec!["2", "e"]);
}

#[test]
fn test_str_isalpha_isnumeric() {
    let script = r#"
print("hello".isalpha())
print("123".isnumeric())
print("café".isalpha())
print("²".isnumeric())
print("abc123".isalpha())
"#;
    assert_eq!(run_python(script), vec!["True", "True", "True", "True", "False"]);
}

#[test]
fn test_unicode_lookup() {
    let script = r#"
import unicodedata
ch = unicodedata.lookup('SNOWMAN')
print(ch)
print(unicodedata.name(ch))
"#;
    assert_eq!(run_python(script), vec!["\u{2603}", "SNOWMAN"]);
}

#[test]
fn test_unicode_decimal_digit() {
    let script = r#"
import unicodedata
print(unicodedata.decimal('0'))
print(unicodedata.digit('9'))
try:
    unicodedata.decimal('A')
except ValueError:
    print("ValueError")
"#;
    assert_eq!(run_python(script), vec!["0", "9", "ValueError"]);
}
