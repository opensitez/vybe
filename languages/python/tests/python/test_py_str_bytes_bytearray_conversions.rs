use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: String, Bytes & Bytearray Conversions — str, bytes, bytearray, memoryview, encodings, mutation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_str_to_bytes_and_back() {
    let src = r#"
s = "Hello World"
b = s.encode("utf-8")
print(type(b).__name__)
print(b)
s_back = b.decode("utf-8")
print(s_back == s)
"#;
    assert_eq!(run_python(src), vec!["bytes", "b'Hello World'", "True"]);
}

#[test]
fn test_py_bytearray_in_place_mutation() {
    let src = r#"
ba = bytearray(b"hello")
ba[0] = 72  # 'H'
ba.extend(b" world")
print(ba.decode())
print(len(ba))
"#;
    assert_eq!(run_python(src), vec!["Hello world", "11"]);
}

#[test]
fn test_py_memoryview_slice_without_copy() {
    let src = r#"
ba = bytearray(b"ABCDEF")
mv = memoryview(ba)
sub = mv[2:5]
print(sub.tobytes().decode())
sub[0] = 90  # 'Z'
print(ba.decode())  # mutated underlying bytearray
"#;
    assert_eq!(run_python(src), vec!["CDE", "ABZDEF"]);
}

#[test]
fn test_py_encoding_error_handlers() {
    let src = r#"
text = "café"
b_strict = text.encode("ascii", errors="ignore")
print(b_strict.decode())

b_replace = text.encode("ascii", errors="replace")
print(b_replace.decode())

b_xml = text.encode("ascii", errors="xmlcharrefreplace")
print(b_xml.decode())
"#;
    assert_eq!(run_python(src), vec!["caf", "caf?", "caf&#233;"]);
}

#[test]
fn test_py_latin1_and_utf16_encodings() {
    let src = r#"
s = "Python"
b_latin = s.encode("latin-1")
b_utf16 = s.encode("utf-16")
print(b_latin.decode("latin-1"))
print(b_utf16.decode("utf-16"))
"#;
    assert_eq!(run_python(src), vec!["Python", "Python"]);
}

#[test]
fn test_py_bytes_formatting_and_hex() {
    let src = r#"
b = bytes.fromhex("48656c6c6f")
print(b.decode())
print(b.hex())
"#;
    assert_eq!(run_python(src), vec!["Hello", "48656c6c6f"]);
}

#[test]
fn test_py_bytearray_methods_split_replace() {
    let src = r#"
ba = bytearray(b"foo,bar,baz")
parts = ba.split(b",")
print([p.decode() for p in parts])

ba.replace(b"bar", b"qux")
print(ba.decode())
"#;
    assert_eq!(
        run_python(src),
        vec!["['foo', 'bar', 'baz']", "foo,bar,baz"]
    );
}

#[test]
fn test_py_bytes_immutability() {
    let src = r#"
b = b"immutable"
try:
    b[0] = 73
except TypeError as e:
    print("TypeError: bytes object does not support item assignment")
"#;
    assert_eq!(
        run_python(src),
        vec!["TypeError: bytes object does not support item assignment"]
    );
}

#[test]
fn test_py_memoryview_cast_format() {
    let src = r#"
import array

a = array.array("i", [1, 2, 3])
mv = memoryview(a)
print(mv.format)
print(mv.itemsize)
print(mv.tolist())
"#;
    assert_eq!(run_python(src), vec!["i", "4", "[1, 2, 3]"]);
}

#[test]
fn test_py_str_maketrans_translate_bytes() {
    let src = r#"
table = bytes.maketrans(b"abc", b"XYZ")
b = b"alphabet"
print(b.translate(table).decode())
"#;
    assert_eq!(run_python(src), vec!["XYZhAbet"]);
}

#[test]
fn test_py_chr_ord_ascii_bounds() {
    let src = r#"
print(ord("A"), ord("z"))
print(chr(65), chr(122))
"#;
    assert_eq!(run_python(src), vec!["65 122", "A z"]);
}

#[test]
fn test_py_bytearray_pop_insert_remove() {
    let src = r#"
ba = bytearray(b"ACD")
ba.insert(1, 66)  # 'B'
print(ba.decode())
pop_val = ba.pop()
print(chr(pop_val))
print(ba.decode())
"#;
    assert_eq!(run_python(src), vec!["ABCD", "D", "ABC"]);
}

#[test]
fn test_py_bytes_join_sequence() {
    let src = r#"
parts = [b"one", b"two", b"three"]
joined = b" | ".join(parts)
print(joined.decode())
"#;
    assert_eq!(run_python(src), vec!["one | two | three"]);
}

#[test]
fn test_py_bytearray_clear_copy() {
    let src = r#"
ba1 = bytearray(b"original")
ba2 = ba1.copy()
ba1.clear()
print(len(ba1))
print(ba2.decode())
"#;
    assert_eq!(run_python(src), vec!["0", "original"]);
}

#[test]
fn test_py_str_isascii_identifier_predicates() {
    let src = r#"
print("ascii_text".isascii())
print("café".isascii())
print("valid_var".isidentifier())
"#;
    assert_eq!(run_python(src), vec!["True", "False", "True"]);
}
