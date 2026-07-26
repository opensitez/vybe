use super::helpers::run_python;

#[test]
fn test_python_base64_basic() {
    let src = r#"
import base64
enc = base64.b64encode(b'abc')
dec = base64.b64decode(enc)
print(enc)
print(dec)
"#;
    assert_eq!(run_python(src), vec!["b'YWJj'", "b'abc'"]);
}

#[test]
fn test_python_binascii_hexlify() {
    let src = r#"
import binascii
hx = binascii.hexlify(b'hi')
out = binascii.unhexlify(hx)
print(hx)
print(out)
"#;
    assert_eq!(run_python(src), vec!["b'6869'", "b'hi'"]);
}
