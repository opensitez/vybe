use super::helpers::run_python;

// string module — Template, Formatter, constants (ascii_letters, digits, punctuation, etc.)

#[test]
fn test_string_ascii_letters_content() {
    let out = run_python(r#"
import string
print(string.ascii_letters[:5])
print(string.ascii_letters[-5:])
"#);
    assert_eq!(out, vec!["abcde", "vwxyz"]);
}

#[test]
fn test_string_ascii_lowercase_and_uppercase() {
    let out = run_python(r#"
import string
print(len(string.ascii_lowercase))
print(len(string.ascii_uppercase))
print(string.ascii_lowercase[0])
print(string.ascii_uppercase[0])
"#);
    assert_eq!(out, vec!["26", "26", "a", "A"]);
}

#[test]
fn test_string_digits_constant() {
    let out = run_python(r#"
import string
print(string.digits)
"#);
    assert_eq!(out, vec!["0123456789"]);
}

#[test]
fn test_string_hexdigits_constant() {
    let out = run_python(r#"
import string
print(string.hexdigits)
"#);
    assert_eq!(out, vec!["0123456789abcdefABCDEF"]);
}

#[test]
fn test_string_octdigits_constant() {
    let out = run_python(r#"
import string
print(string.octdigits)
"#);
    assert_eq!(out, vec!["01234567"]);
}

#[test]
fn test_string_punctuation_contains_expected() {
    let out = run_python(r#"
import string
print("!" in string.punctuation)
print("@" in string.punctuation)
print("a" in string.punctuation)
"#);
    assert_eq!(out, vec!["True", "True", "False"]);
}

#[test]
fn test_string_whitespace_contains_space_tab_newline() {
    let out = run_python(r#"
import string
print(" " in string.whitespace)
print("\t" in string.whitespace)
print("\n" in string.whitespace)
"#);
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_string_printable_length() {
    let out = run_python(r#"
import string
print(len(string.printable))
"#);
    assert_eq!(out, vec!["100"]);
}

#[test]
fn test_string_template_substitute() {
    let out = run_python(r#"
import string
t = string.Template("Hello, $name!")
print(t.substitute(name="World"))
"#);
    assert_eq!(out, vec!["Hello, World!"]);
}

#[test]
fn test_string_template_substitute_with_braces() {
    let out = run_python(r#"
import string
t = string.Template("${item}s are available")
print(t.substitute(item="book"))
"#);
    assert_eq!(out, vec!["books are available"]);
}

#[test]
fn test_string_template_safe_substitute_missing_key() {
    let out = run_python(r#"
import string
t = string.Template("Hello $name, you have $count messages")
print(t.safe_substitute(name="Alice"))
"#);
    assert_eq!(out, vec!["Hello Alice, you have $count messages"]);
}

#[test]
fn test_string_template_substitute_missing_key_raises() {
    let out = run_python(r#"
import string
t = string.Template("$missing")
try:
    t.substitute({})
except KeyError:
    print("KeyError")
"#);
    assert_eq!(out, vec!["KeyError"]);
}

#[test]
fn test_string_template_dict_argument() {
    let out = run_python(r#"
import string
t = string.Template("$x + $y = $z")
print(t.substitute({"x": "1", "y": "2", "z": "3"}))
"#);
    assert_eq!(out, vec!["1 + 2 = 3"]);
}

#[test]
fn test_string_template_custom_delimiter() {
    let out = run_python(r#"
import string
class MyTemplate(string.Template):
    delimiter = "!"
t = MyTemplate("Hello !name!")
print(t.substitute(name="World"))
"#);
    assert_eq!(out, vec!["Hello World!"]);
}

#[test]
fn test_string_formatter_vformat() {
    let out = run_python(r#"
import string
f = string.Formatter()
result = f.format("{0} and {1}", "alpha", "beta")
print(result)
"#);
    assert_eq!(out, vec!["alpha and beta"]);
}

#[test]
fn test_string_formatter_named_fields() {
    let out = run_python(r#"
import string
f = string.Formatter()
print(f.format("{name} is {age}", name="Alice", age=30))
"#);
    assert_eq!(out, vec!["Alice is 30"]);
}

#[test]
fn test_string_capwords() {
    let out = run_python(r#"
import string
print(string.capwords("hello world foo"))
"#);
    assert_eq!(out, vec!["Hello World Foo"]);
}

#[test]
fn test_string_capwords_custom_separator() {
    let out = run_python(r#"
import string
print(string.capwords("hello-world-foo", sep="-"))
"#);
    assert_eq!(out, vec!["Hello-World-Foo"]);
}

#[test]
fn test_string_template_get_identifiers() {
    let out = run_python(r#"
import string
t = string.Template("$foo and $bar and $foo")
ids = t.get_identifiers()
print(sorted(ids))
"#);
    assert_eq!(out, vec!["['bar', 'foo']"]);
}

#[test]
fn test_string_template_is_valid_true() {
    let out = run_python(r#"
import string
t = string.Template("$name")
print(t.is_valid())
"#);
    assert_eq!(out, vec!["True"]);
}
