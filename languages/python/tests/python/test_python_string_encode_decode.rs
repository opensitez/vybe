// Python string encode/decode — utf-8, latin-1, errors, codecs
use super::helpers::run_python;

#[test]
fn test_encode_utf8() {
    let script = r#"
s = "hello"
b = s.encode('utf-8')
print(b)
print(type(b).__name__)
"#;
    assert_eq!(run_python(script), vec!["b'hello'", "bytes"]);
}

#[test]
fn test_encode_latin1() {
    let script = r#"
s = "caf\u00e9"
b = s.encode('latin-1')
print(list(b))
"#;
    assert_eq!(run_python(script), vec!["[99, 97, 102, 233]"]);
}

#[test]
fn test_decode_utf8() {
    let script = r#"
b = b"world"
s = b.decode('utf-8')
print(s)
print(type(s).__name__)
"#;
    assert_eq!(run_python(script), vec!["world", "str"]);
}

#[test]
fn test_encode_errors_ignore() {
    let script = r#"
s = "hello \u00e9 world"
b = s.encode('ascii', errors='ignore')
print(b.decode('ascii'))
"#;
    assert_eq!(run_python(script), vec!["hello  world"]);
}

#[test]
fn test_encode_errors_replace() {
    let script = r#"
s = "caf\u00e9"
b = s.encode('ascii', errors='replace')
print(b.decode('ascii'))
"#;
    assert_eq!(run_python(script), vec!["caf?"]);
}

#[test]
fn test_encode_decode_roundtrip() {
    let script = r#"
original = "Hello, 世界!"
encoded = original.encode('utf-8')
decoded = encoded.decode('utf-8')
print(original == decoded)
print(len(encoded))
"#;
    assert_eq!(run_python(script), vec!["True", "14"]);
}

#[test]
fn test_decode_errors_strict() {
    let script = r#"
b = bytes([0xFF, 0xFE])
try:
    b.decode('utf-8', errors='strict')
    print("no_error")
except UnicodeDecodeError:
    print("UnicodeDecodeError")
"#;
    assert_eq!(run_python(script), vec!["UnicodeDecodeError"]);
}

#[test]
fn test_string_encode_ascii() {
    let script = r#"
s = "ABC"
b = s.encode('ascii')
print(list(b))
"#;
    assert_eq!(run_python(script), vec!["[65, 66, 67]"]);
}
