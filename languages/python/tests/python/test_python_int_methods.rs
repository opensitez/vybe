// Python int methods — bit_length, to_bytes, from_bytes, as_integer_ratio
use super::helpers::run_python;

#[test]
fn test_int_bit_length() {
    let script = r#"
print((0).bit_length())
print((1).bit_length())
print((255).bit_length())
print((256).bit_length())
"#;
    assert_eq!(run_python(script), vec!["0", "1", "8", "9"]);
}

#[test]
fn test_int_bit_count() {
    let script = r#"
print((0b1010).bit_count())
print((0b1111).bit_count())
print((0).bit_count())
"#;
    assert_eq!(run_python(script), vec!["2", "4", "0"]);
}

#[test]
fn test_int_to_bytes() {
    let script = r#"
n = 1024
b = n.to_bytes(2, byteorder='big')
print(list(b))
"#;
    assert_eq!(run_python(script), vec!["[4, 0]"]);
}

#[test]
fn test_int_from_bytes() {
    let script = r#"
b = bytes([4, 0])
n = int.from_bytes(b, byteorder='big')
print(n)
"#;
    assert_eq!(run_python(script), vec!["1024"]);
}

#[test]
fn test_int_from_bytes_little_endian() {
    let script = r#"
b = bytes([1, 0])
n = int.from_bytes(b, byteorder='little')
print(n)
"#;
    assert_eq!(run_python(script), vec!["1"]);
}

#[test]
fn test_int_as_integer_ratio() {
    let script = r#"
print((10).as_integer_ratio())
print((-3).as_integer_ratio())
"#;
    assert_eq!(run_python(script), vec!["(10, 1)", "(-3, 1)"]);
}

#[test]
fn test_int_is_integer() {
    let script = r#"
print((42).is_integer())
"#;
    assert_eq!(run_python(script), vec!["True"]);
}

#[test]
fn test_int_conjugate_real_imag() {
    let script = r#"
n = 7
print(n.conjugate())
print(n.real)
print(n.imag)
"#;
    assert_eq!(run_python(script), vec!["7", "7", "0"]);
}

#[test]
fn test_int_numerator_denominator() {
    let script = r#"
n = -5
print(n.numerator)
print(n.denominator)
"#;
    assert_eq!(run_python(script), vec!["-5", "1"]);
}
