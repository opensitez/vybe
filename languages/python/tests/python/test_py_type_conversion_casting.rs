use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Type Conversion & Casting — int, float, str, bool, list, dict, set, tuple, bytes, bytearray, complex
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_int_conversions_bases() {
    let src = r#"
print(int("123"))
print(int("0b1010", 2))
print(int("0o755", 8))
print(int("0x1a3f", 16))
print(int(3.99))
print(int(-3.99))
"#;
    assert_eq!(run_python(src), vec!["123", "10", "493", "6719", "3", "-3"]);
}

#[test]
fn test_py_float_conversions_strings() {
    let src = r#"
print(float("123.45"))
print(float("-1e-3"))
print(float("inf"))
print(float("-inf"))
print(float("nan") != float("nan"))
"#;
    assert_eq!(
        run_python(src),
        vec!["123.45", "-0.001", "inf", "-inf", "True"]
    );
}

#[test]
fn test_py_bool_truthiness_casting() {
    let src = r#"
print(bool(1))
print(bool(0))
print(bool("hello"))
print(bool(""))
print(bool([0]))
print(bool([]))
print(bool(None))
"#;
    assert_eq!(
        run_python(src),
        vec!["True", "False", "True", "False", "True", "False", "False"]
    );
}

#[test]
fn test_py_container_conversions() {
    let src = r#"
s = "hello"
print(list(s))
print(tuple(s))
print(sorted(set(s)))
pairs = [("a", 1), ("b", 2)]
print(dict(pairs))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "['h', 'e', 'l', 'l', 'o']",
            "('h', 'e', 'l', 'l', 'o')",
            "['e', 'h', 'l', 'o']",
            "{'a': 1, 'b': 2}"
        ]
    );
}

#[test]
fn test_py_bytes_bytearray_conversions() {
    let src = r#"
b1 = bytes([65, 66, 67])
print(b1)
ba = bytearray(b1)
ba[0] = 90
print(ba.decode())
print(bytes(ba))
"#;
    assert_eq!(run_python(src), vec!["b'ABC'", "ZBC", "b'ZBC'"]);
}

#[test]
fn test_py_complex_number_conversions() {
    let src = r#"
c1 = complex(3, 4)
print(c1)
c2 = complex("1+2j")
print(c2)
print(c1.real, c1.imag)
"#;
    assert_eq!(run_python(src), vec!["(3+4j)", "(1+2j)", "3.0 4.0"]);
}

#[test]
fn test_py_custom_int_trunc_index_dunders() {
    let src = r#"
class CustomIndex:
    def __index__(self):
        return 4

class CustomTrunc:
    def __trunc__(self):
        return 99

lst = [10, 20, 30, 40, 50]
print(lst[CustomIndex()])
print(int(CustomTrunc()))
"#;
    assert_eq!(run_python(src), vec!["50", "99"]);
}

#[test]
fn test_py_custom_float_dunder() {
    let src = r#"
class Distance:
    def __init__(self, meters):
        self.meters = meters

    def __float__(self):
        return float(self.meters)

d = Distance(12.5)
print(float(d))
print(f"{float(d):.2f}")
"#;
    assert_eq!(run_python(src), vec!["12.5", "12.50"]);
}

#[test]
fn test_py_hex_oct_bin_builtins() {
    let src = r#"
val = 255
print(hex(val))
print(oct(val))
print(bin(val))
"#;
    assert_eq!(run_python(src), vec!["0xff", "0o377", "0b11111111"]);
}

#[test]
fn test_py_chr_ord_unicode_code_point() {
    let src = r#"
ch = 'A'
code = ord(ch)
print(code)
print(chr(code))
print(chr(0x1F600))  # 😀
"#;
    assert_eq!(run_python(src), vec!["65", "A", "😀"]);
}
