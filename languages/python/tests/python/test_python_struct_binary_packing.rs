use super::helpers::run_python;

#[test]
fn test_python_struct_pack_unpack_int() {
    let src = r#"
import struct
blob = struct.pack('>I', 1024)
print(blob)
print(struct.unpack('>I', blob)[0])
"#;
    assert_eq!(run_python(src), vec!["b'\x00\x00\x04\x00'", "1024"]);
}

#[test]
fn test_python_struct_calcsize_endian() {
    let src = r#"
import struct
print(struct.calcsize('<h'))
print(struct.calcsize('!f'))
"#;
    assert_eq!(run_python(src), vec!["2", "4"]);
}
