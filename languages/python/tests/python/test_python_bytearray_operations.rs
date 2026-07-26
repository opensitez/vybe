// Python bytearray — mutable bytes, mutation, slicing, methods
use super::helpers::run_python;

#[test]
fn test_bytearray_create_from_string() {
    let script = r#"
ba = bytearray("hello", "utf-8")
print(len(ba))
print(ba[0])
"#;
    assert_eq!(run_python(script), vec!["5", "104"]);
}

#[test]
fn test_bytearray_mutate_item() {
    let script = r#"
ba = bytearray(b"abc")
ba[0] = 65
print(bytes(ba))
"#;
    assert_eq!(run_python(script), vec!["b'Abc'"]);
}

#[test]
fn test_bytearray_append_extend() {
    let script = r#"
ba = bytearray(b"AB")
ba.append(67)
ba.extend(b"DE")
print(bytes(ba))
"#;
    assert_eq!(run_python(script), vec!["b'ABCDE'"]);
}

#[test]
fn test_bytearray_pop() {
    let script = r#"
ba = bytearray(b"XYZ")
v = ba.pop()
print(v)
print(bytes(ba))
"#;
    assert_eq!(run_python(script), vec!["90", "b'XY'"]);
}

#[test]
fn test_bytearray_insert() {
    let script = r#"
ba = bytearray(b"AC")
ba.insert(1, 66)
print(bytes(ba))
"#;
    assert_eq!(run_python(script), vec!["b'ABC'"]);
}

#[test]
fn test_bytearray_slice_assignment() {
    let script = r#"
ba = bytearray(b"hello")
ba[1:4] = b"ELL"
print(bytes(ba))
"#;
    assert_eq!(run_python(script), vec!["b'hELLo'"]);
}

#[test]
fn test_bytearray_decode() {
    let script = r#"
ba = bytearray(b"world")
print(ba.decode("utf-8"))
"#;
    assert_eq!(run_python(script), vec!["world"]);
}

#[test]
fn test_bytearray_hex_fromhex() {
    let script = r#"
ba = bytearray.fromhex("414243")
print(bytes(ba))
print(ba.hex())
"#;
    assert_eq!(run_python(script), vec!["b'ABC'", "414243"]);
}

#[test]
fn test_bytearray_replace() {
    let script = r#"
ba = bytearray(b"aabbcc")
result = ba.replace(b"bb", b"XX")
print(bytes(result))
"#;
    assert_eq!(run_python(script), vec!["b'aaXXcc'"]);
}

#[test]
fn test_bytearray_find() {
    let script = r#"
ba = bytearray(b"hello world")
print(ba.find(b"world"))
print(ba.find(b"xyz"))
"#;
    assert_eq!(run_python(script), vec!["6", "-1"]);
}
