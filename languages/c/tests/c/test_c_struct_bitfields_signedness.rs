use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn bitfield_sign_signed_int() {
    assert_eq!(
        run_c(
            "struct S { signed int a:3; }; int main() { struct S s; s.a = 7; /* 7 is 111, in 3-bit signed it's -1 */ printf(\"%d\", s.a); return 0; }"
        ),
        vec!["-1"]
    );
}
#[test]
fn bitfield_sign_unsigned_int() {
    assert_eq!(
        run_c(
            "struct S { unsigned int a:3; }; int main() { struct S s; s.a = 7; printf(\"%d\", s.a); return 0; }"
        ),
        vec!["7"]
    );
}
#[test]
fn bitfield_sign_plain_int_impl_defined() {
    assert_eq!(
        run_c(
            "struct S { int a:3; }; int main() { struct S s; s.a = 7; int val = s.a; printf(\"%d\", val == 7 || val == -1); return 0; }"
        ),
        vec!["1"]
    );
} // Implementation defined if 'int' bitfield is signed
#[test]
fn bitfield_sign_char_signed() {
    assert_eq!(
        run_c(
            "struct S { signed char a:2; }; int main() { struct S s; s.a = 3; /* 11 -> -1 */ printf(\"%d\", s.a); return 0; }"
        ),
        vec!["-1"]
    );
}
#[test]
fn bitfield_sign_char_unsigned() {
    assert_eq!(
        run_c(
            "struct S { unsigned char a:2; }; int main() { struct S s; s.a = 3; printf(\"%d\", s.a); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn bitfield_sign_negative_assignment() {
    assert_eq!(
        run_c(
            "struct S { signed int a:4; }; int main() { struct S s; s.a = -2; printf(\"%d\", s.a); return 0; }"
        ),
        vec!["-2"]
    );
}
#[test]
fn bitfield_sign_negative_unsigned_assignment() {
    assert_eq!(
        run_c(
            "struct S { unsigned int a:4; }; int main() { struct S s; s.a = -2; /* Wraps around modulo 16 -> 14 */ printf(\"%d\", s.a); return 0; }"
        ),
        vec!["14"]
    );
}
#[test]
fn bitfield_sign_one_bit_signed() {
    assert_eq!(
        run_c(
            "struct S { signed int a:1; }; int main() { struct S s; s.a = 1; /* 1 bit signed can only hold 0 or -1 */ printf(\"%d\", s.a); return 0; }"
        ),
        vec!["-1"]
    );
}
#[test]
fn bitfield_sign_one_bit_unsigned() {
    assert_eq!(
        run_c(
            "struct S { unsigned int a:1; }; int main() { struct S s; s.a = 1; printf(\"%d\", s.a); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn bitfield_sign_promotion_signed() {
    assert_eq!(
        run_c(
            "struct S { signed int a:4; }; int main() { struct S s = {-1}; printf(\"%d\", s.a + 1); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn bitfield_sign_promotion_unsigned_to_signed() {
    assert_eq!(
        run_c(
            "struct S { unsigned int a:4; }; int main() { struct S s = {15}; printf(\"%d\", s.a + 1); return 0; }"
        ),
        vec!["16"]
    );
} // 4-bit unsigned fits in int, promotes to signed int
#[test]
fn bitfield_sign_promotion_unsigned_to_unsigned() {
    assert_eq!(
        run_c(
            "struct S { unsigned int a:32; }; int main() { struct S s = {4294967295U}; printf(\"%u\", s.a + 1); return 0; }"
        ),
        vec!["0"]
    );
} // Assuming 32-bit int, fits unsigned int, promotes to unsigned int
#[test]
fn bitfield_sign_comparisons() {
    assert_eq!(
        run_c(
            "struct S { signed int a:3; }; int main() { struct S s = {7}; /* -1 */ printf(\"%d\", s.a < 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn bitfield_sign_bitwise_not() {
    assert_eq!(
        run_c(
            "struct S { unsigned int a:3; }; int main() { struct S s = {5}; /* ~5 is ~0b101 -> ...1111010 -> -6 in int */ printf(\"%d\", ~s.a); return 0; }"
        ),
        vec!["-6"]
    );
}
#[test]
fn bitfield_sign_shift() {
    assert_eq!(
        run_c(
            "struct S { unsigned int a:4; }; int main() { struct S s = {3}; printf(\"%d\", s.a << 2); return 0; }"
        ),
        vec!["12"]
    );
}
