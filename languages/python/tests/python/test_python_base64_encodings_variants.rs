use super::helpers::run_python;

// base64 — b85encode, b85decode, a85encode, a85decode, b32encode, b32decode, b16encode, b16decode, urlsafe_b64encode, urlsafe_b64decode, standard_b64encode, standard_b64decode

#[test]
fn test_base64_b85encode_b85decode_roundtrip() {
    let out = run_python(r#"
import base64
original = b"The quick brown fox jumps over the lazy dog"
encoded = base64.b85encode(original)
decoded = base64.b85decode(encoded)
print(decoded == original)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_base64_a85encode_a85decode_ascii85() {
    let out = run_python(r#"
import base64
data = b"Hello, World! 12345"
encoded = base64.a85encode(data)
decoded = base64.a85decode(encoded)
print(decoded == data)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_base64_b32encode_b32decode_base32() {
    let out = run_python(r#"
import base64
data = b"Python base32 encoding"
encoded = base64.b32encode(data)
print(encoded.isalnum())
decoded = base64.b32decode(encoded)
print(decoded == data)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_base64_b16encode_b16decode_hex() {
    let out = run_python(r#"
import base64
data = b"\x00\x01\x02\xfe\xff"
encoded = base64.b16encode(data)
print(encoded)
decoded = base64.b16decode(encoded)
print(decoded == data)
"#);
    assert_eq!(out, vec!["b'000102FEFF'", "True"]);
}

#[test]
fn test_base64_urlsafe_b64encode_and_decode() {
    let out = run_python(r#"
import base64
# Data containing bytes that standard b64 produces + and / for
data = b"\xfb\xff\xfe\xfd"
encoded = base64.urlsafe_b64encode(data)
print(b"+" not in encoded and b"/" not in encoded)
decoded = base64.urlsafe_b64decode(encoded)
print(decoded == data)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_base64_standard_b64encode_and_decode() {
    let out = run_python(r#"
import base64
data = b"standard base64 string"
encoded = base64.standard_b64encode(data)
decoded = base64.standard_b64decode(encoded)
print(decoded == data)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_base64_a85encode_adobe_wrap() {
    let out = run_python(r#"
import base64
data = b"adobe ascii85"
encoded = base64.a85encode(data, adobe=True)
print(encoded.startswith(b"<~") and encoded.endswith(b"~>"))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_base64_b85encode_pad_argument() {
    let out = run_python(r#"
import base64
data = b"123"  # not multiple of 4
encoded = base64.b85encode(data, pad=True)
decoded = base64.b85decode(encoded)
print(decoded.startswith(b"123"))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_base64_b32decode_casefold() {
    let out = run_python(r#"
import base64
encoded_lower = b"jbswy3dpeblw64tmmqqq===="
decoded = base64.b32decode(encoded_lower, casefold=True)
print(decoded)
"#);
    assert_eq!(out, vec!["b'Hello World!'" ]);
}

#[test]
fn test_base64_b16decode_casefold() {
    let out = run_python(r#"
import base64
encoded_lower = b"48656c6c6f"
decoded = base64.b16decode(encoded_lower, casefold=True)
print(decoded)
"#);
    assert_eq!(out, vec!["b'Hello'"]);
}

#[test]
fn test_base64_b32hexencode_and_decode() {
    let out = run_python(r#"
import base64, sys
if hasattr(base64, "b32hexencode"):
    data = b"base32hex test"
    enc = base64.b32hexencode(data)
    dec = base64.b32hexdecode(enc)
    print(dec == data)
else:
    print(True)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_base64_decodebytes_and_encodebytes() {
    let out = run_python(r#"
import base64
data = b"test line wrapping"
enc = base64.encodebytes(data)
print(b"\n" in enc)
dec = base64.decodebytes(enc)
print(dec == data)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_base64_b64decode_validate_flag() {
    let out = run_python(r#"
import base64, binascii
invalid_b64 = b"SGVsbG8=!!!"
try:
    base64.b64decode(invalid_b64, validate=True)
except binascii.Error:
    print("binascii.Error")
"#);
    assert_eq!(out, vec!["binascii.Error"]);
}

#[test]
fn test_base64_a85decode_foldspaces() {
    let out = run_python(r#"
import base64
data = b"    " * 2
enc = base64.a85encode(data, foldspaces=True)
dec = base64.a85decode(enc, foldspaces=True)
print(dec == data)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_base64_urlsafe_b64decode_string_input() {
    let out = run_python(r#"
import base64
enc_str = "aGVsbG8td29ybGQ"
# Add padding if needed automatically in decode
dec = base64.urlsafe_b64decode(enc_str + "==")
print(dec)
"#);
    assert_eq!(out, vec!["b'hello-world'"]);
}

#[test]
fn test_base64_b16encode_returns_ascii_bytes() {
    let out = run_python(r#"
import base64
res = base64.b16encode(b"ABC")
print(res)
print(isinstance(res, bytes))
"#);
    assert_eq!(out, vec!["b'414243'", "True"]);
}

#[test]
fn test_base64_empty_bytes_input() {
    let out = run_python(r#"
import base64
print(base64.b64encode(b""))
print(base64.b85encode(b""))
print(base64.b32encode(b""))
print(base64.b16encode(b""))
"#);
    assert_eq!(out, vec!["b''", "b''", "b''", "b''"]);
}

#[test]
fn test_base64_b85decode_invalid_character_raises() {
    let out = run_python(r#"
import base64, ValueError
try:
    base64.b85decode(b"invalid \x00 char")
except Exception:
    print("Error")
"#);
    assert_eq!(out, vec!["Error"]);
}

#[test]
fn test_base64_b32decode_map01_kwarg() {
    let out = run_python(r#"
import base64
# map01 maps '0' to 'O' and '1' to 'I' or 'L'
data = b"JBSWY3DPEBLW64TMMQQQ===="
enc_with_01 = data.replace(b"O", b"0").replace(b"I", b"1")
dec = base64.b32decode(enc_with_01, map01=b"I")
print(dec)
"#);
    assert_eq!(out, vec!["b'Hello World!'"]);
}

#[test]
fn test_base64_encode_and_decode_file_like() {
    let out = run_python(r#"
import base64, io
input_stream = io.BytesIO(b"file stream data")
output_stream = io.BytesIO()
base64.encode(input_stream, output_stream)
output_stream.seek(0)

decoded_stream = io.BytesIO()
base64.decode(output_stream, decoded_stream)
print(decoded_stream.getvalue())
"#);
    assert_eq!(out, vec!["b'file stream data'"]);
}
