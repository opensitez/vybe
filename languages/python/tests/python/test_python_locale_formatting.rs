// Python locale — getlocale, setlocale, format_string, currency, number grouping
use super::helpers::run_python;

#[test]
fn test_locale_getlocale() {
    let script = r#"
import locale
loc = locale.getlocale()
print(isinstance(loc, tuple))
print(len(loc) == 2)
"#;
    assert_eq!(run_python(script), vec!["True", "True"]);
}

#[test]
fn test_locale_setlocale_c() {
    let script = r#"
import locale
locale.setlocale(locale.LC_ALL, 'C')
result = locale.getlocale(locale.LC_ALL)
print(result[0] in ('C', None, 'C.UTF-8') or result[0] is None or 'C' in str(result))
"#;
    assert_eq!(run_python(script), vec!["True"]);
}

#[test]
fn test_locale_atof() {
    let script = r#"
import locale
locale.setlocale(locale.LC_ALL, 'C')
print(locale.atof('3.14'))
print(locale.atoi('42'))
"#;
    assert_eq!(run_python(script), vec!["3.14", "42"]);
}

#[test]
fn test_locale_str() {
    let script = r#"
import locale
locale.setlocale(locale.LC_ALL, 'C')
print(locale.str(3.14))
"#;
    assert_eq!(run_python(script), vec!["3.14"]);
}

#[test]
fn test_locale_constants_exist() {
    let script = r#"
import locale
print(hasattr(locale, 'LC_ALL'))
print(hasattr(locale, 'LC_TIME'))
print(hasattr(locale, 'LC_NUMERIC'))
print(hasattr(locale, 'LC_MONETARY'))
"#;
    assert_eq!(run_python(script), vec!["True", "True", "True", "True"]);
}

#[test]
fn test_locale_format_string() {
    let script = r#"
import locale
locale.setlocale(locale.LC_ALL, 'C')
result = locale.format_string("%.2f", 1234.567)
print(result)
"#;
    assert_eq!(run_python(script), vec!["1234.57"]);
}
