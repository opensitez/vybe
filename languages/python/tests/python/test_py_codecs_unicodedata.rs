use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: codecs + unicodedata + base64 + binascii — text encodings, unicode normalization, base64, hex
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_base64_encode_decode() {
    let src = r#"
import base64

raw = b"Hello, World!"
b64 = base64.b64encode(raw)
print(b64.decode())

restored = base64.b64decode(b64)
print(restored == raw)
"#;
    assert_eq!(run_python(src), vec!["SGVsbG8sIFdvcmxkIQ==", "True"]);
}

#[test]
fn test_py_base64_urlsafe_b64encode() {
    let src = r#"
import base64

raw = b"\xfb\xff\xfe"
urlsafe = base64.urlsafe_b64encode(raw)
print("+" not in urlsafe.decode())
print("/" not in urlsafe.decode())
print(base64.urlsafe_b64decode(urlsafe) == raw)
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_unicodedata_name_and_lookup() {
    let src = r#"
import unicodedata

print(unicodedata.name("A"))
print(unicodedata.name("€"))
print(unicodedata.lookup("EURO SIGN"))
"#;
    assert_eq!(
        run_python(src),
        vec!["LATIN CAPITAL LETTER A", "EURO SIGN", "€"]
    );
}

#[test]
fn test_py_unicodedata_normalize_nfc_nfd() {
    let src = r#"
import unicodedata

# 'é' as single char vs 'e' + combining acute accent
single = "\u00e9"
decomposed = "e\u0301"

print(len(single))
print(len(decomposed))
print(single != decomposed)

nfc1 = unicodedata.normalize("NFC", single)
nfc2 = unicodedata.normalize("NFC", decomposed)
print(nfc1 == nfc2)
print(len(nfc1))

nfd1 = unicodedata.normalize("NFD", single)
print(len(nfd1))
"#;
    assert_eq!(run_python(src), vec!["1", "2", "True", "True", "1", "2"]);
}

#[test]
fn test_py_unicodedata_category_numeric() {
    let src = r#"
import unicodedata

print(unicodedata.category("A"))  # Lu = Letter, uppercase
print(unicodedata.category("1"))  # Nd = Number, decimal digit
print(unicodedata.category(" "))  # Zs = Separator, space

print(unicodedata.numeric("5"))
print(unicodedata.numeric("½"))
"#;
    assert_eq!(run_python(src), vec!["Lu", "Nd", "Zs", "5.0", "0.5"]);
}

#[test]
fn test_py_binascii_hexlify_unhexlify() {
    let src = r#"
import binascii

data = b"\x00\xff\x10\x20"
hex_str = binascii.hexlify(data)
print(hex_str.decode())

restored = binascii.unhexlify(hex_str)
print(restored == data)
"#;
    assert_eq!(run_python(src), vec!["00ff1020", "True"]);
}

#[test]
fn test_py_codecs_encode_decode_rot13() {
    let src = r#"
import codecs

text = "Hello World"
rot13 = codecs.encode(text, "rot_13")
print(rot13)
print(codecs.decode(rot13, "rot_13") == text)
"#;
    assert_eq!(run_python(src), vec!["Uuryyb Jbeyq", "True"]);
}

#[test]
fn test_py_codecs_open_file_encoding() {
    let src = r#"
import codecs, tempfile, os

with tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as f:
    fname = f.name

with codecs.open(fname, "w", encoding="utf-8") as f:
    f.write("こんにちは世界")

with codecs.open(fname, "r", encoding="utf-8") as f:
    content = f.read()

os.unlink(fname)
print(content)
"#;
    assert_eq!(run_python(src), vec!["こんにちは世界"]);
}

#[test]
fn test_py_base64_b32_b16_encode() {
    let src = r#"
import base64

raw = b"foo"
b32 = base64.b32encode(raw)
print(b32.decode())
print(base64.b32decode(b32) == raw)

b16 = base64.b16encode(raw)
print(b16.decode())
print(base64.b16decode(b16) == raw)
"#;
    assert_eq!(run_python(src), vec!["MZXW6===", "True", "666F6F", "True"]);
}

#[test]
fn test_py_unicodedata_east_asian_width() {
    let src = r#"
import unicodedata

print(unicodedata.east_asian_width("A"))
print(unicodedata.east_asian_width("漢"))
"#;
    assert_eq!(run_python(src), vec!["Na", "W"]);
}
