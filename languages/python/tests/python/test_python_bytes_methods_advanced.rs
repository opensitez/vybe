use super::helpers::run_python;

#[test]
fn test_python_bytes_hex_and_fromhex() {
    let src = r#"
b = b'abc'
print(b.hex())
print(bytes.fromhex('61 62 63'))
"#;
    assert_eq!(run_python(src), vec!["616263", "b'abc'"]);
}

#[test]
fn test_python_bytes_methods_split_replace() {
    let src = r#"
b = b'one,two,three'
print(b.split(b','))
print(b.replace(b',', b';', 1))
"#;
    assert_eq!(
        run_python(src),
        vec!["[b'one', b'two', b'three']", "b'one;two,three'"]
    );
}
