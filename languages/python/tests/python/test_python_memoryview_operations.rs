// Python memoryview — buffer protocol, slicing, cast, formats
use super::helpers::run_python;

#[test]
fn test_memoryview_from_bytes() {
    let script = r#"
mv = memoryview(b"hello")
print(mv[0])
print(bytes(mv[1:4]))
"#;
    assert_eq!(run_python(script), vec!["104", "b'ell'"]);
}

#[test]
fn test_memoryview_from_bytearray() {
    let script = r#"
ba = bytearray(b"ABCDE")
mv = memoryview(ba)
mv[1] = 88  # 'X'
print(bytes(ba))
"#;
    assert_eq!(run_python(script), vec!["b'AXCDE'"]);
}

#[test]
fn test_memoryview_format_and_itemsize() {
    let script = r#"
mv = memoryview(b"test")
print(mv.format)
print(mv.itemsize)
print(mv.nbytes)
"#;
    assert_eq!(run_python(script), vec!["B", "1", "4"]);
}

#[test]
fn test_memoryview_shape_ndim() {
    let script = r#"
mv = memoryview(b"12345678")
print(mv.ndim)
print(mv.shape)
"#;
    assert_eq!(run_python(script), vec!["1", "(8,)"]);
}

#[test]
fn test_memoryview_tobytes() {
    let script = r#"
mv = memoryview(b"data")
print(mv.tobytes())
"#;
    assert_eq!(run_python(script), vec!["b'data'"]);
}

#[test]
fn test_memoryview_tolist() {
    let script = r#"
mv = memoryview(b"\x01\x02\x03")
print(mv.tolist())
"#;
    assert_eq!(run_python(script), vec!["[1, 2, 3]"]);
}

#[test]
fn test_memoryview_cast() {
    let script = r#"
import array
a = array.array('h', [1, 2, 3, 4])
mv = memoryview(a).cast('B')
print(len(mv))  # 4 shorts = 8 bytes
"#;
    assert_eq!(run_python(script), vec!["8"]);
}
