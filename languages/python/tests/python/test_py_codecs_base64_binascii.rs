use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Codecs, Base64 & Binascii Encodings — base64, binascii, hex, urlsafe, codecs.encode, codecs.decode
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_base64_encode_decode_roundtrip() {
    let src = r#"
import base64

data = b"Python 3.12 Web App"
encoded = base64.b64encode(data)
print(encoded.decode("ascii"))

decoded = base64.b64decode(encoded)
print(decoded == data)
"#;
    assert_eq!(
        run_python(src),
        vec!["UHl0aG9uIDMuMTIgV2ViIEFwcA==", "True"]
    );
}

#[test]
fn test_py_base64_urlsafe_encode_decode() {
    let src = r#"
import base64

raw = b"\xff\xe0\x00\x10JFIF"
urlsafe = base64.urlsafe_b64encode(raw)
print(urlsafe.decode())
print("+" not in urlsafe.decode())
print("/" not in urlsafe.decode())
"#;
    assert_eq!(run_python(src), vec!["_-AAEEpGSUY=", "True", "True"]);
}

#[test]
fn test_py_binascii_hexlify_unhexlify_data() {
    let src = r#"
import binascii

data = b"DEADBEEF\x00\x01"
hex_repr = binascii.hexlify(data)
print(hex_repr.decode())

restored = binascii.unhexlify(hex_repr)
print(restored == data)
"#;
    assert_eq!(run_python(src), vec!["44454144424545460001", "True"]);
}

#[test]
fn test_py_binascii_crc32_checksum() {
    let src = r#"
import binascii

crc = binascii.crc32(b"123456789")
print(crc)
"#;
    assert_eq!(run_python(src), vec!["3421780262"]);
}

#[test]
fn test_py_codecs_rot13_cipher() {
    let src = r#"
import codecs

msg = "Hello World 2024"
encrypted = codecs.encode(msg, "rot_13")
print(encrypted)

decrypted = codecs.decode(encrypted, "rot_13")
print(decrypted)
"#;
    assert_eq!(
        run_python(src),
        vec!["Uuryyb Jbeyq 2024", "Hello World 2024"]
    );
}

#[test]
fn test_py_codecs_hex_encode_decode() {
    let src = r#"
import codecs

data = b"abc"
hex_str = codecs.encode(data, "hex")
print(hex_str.decode())
print(codecs.decode(hex_str, "hex") == data)
"#;
    assert_eq!(run_python(src), vec!["616263", "True"]);
}

#[test]
fn test_py_base64_b32_b16_encodings() {
    let src = r#"
import base64

data = b"test"
b32 = base64.b32encode(data)
print(b32.decode())
print(base64.b32decode(b32) == data)

b16 = base64.b16encode(data)
print(b16.decode())
print(base64.b16decode(b16) == data)
"#;
    assert_eq!(
        run_python(src),
        vec!["ORSXG5A=", "True", "74657374", "True"]
    );
}

#[test]
fn test_py_codecs_lookup_info() {
    let src = r#"
import codecs

info = codecs.lookup("utf-8")
print(info.name)
"#;
    assert_eq!(run_python(src), vec!["utf-8"]);
}

#[test]
fn test_py_binascii_a2b_base64_b2a_base64() {
    let src = r#"
import binascii

data = b"hello world"
b64_line = binascii.b2a_base64(data)
print(b64_line.decode().strip())
print(binascii.a2b_base64(b64_line) == data)
"#;
    assert_eq!(run_python(src), vec!["aGVsbG8gd29ybGQ=", "True"]);
}

#[test]
fn test_py_codecs_error_handlers_ignore_replace() {
    let src = r#"
import codecs

bad_bytes = b"hello \xff world"
print(bad_bytes.decode("ascii", errors="ignore"))
print(bad_bytes.decode("ascii", errors="replace"))
"#;
    assert_eq!(run_python(src), vec!["hello  world", "hello  world"]);
}
