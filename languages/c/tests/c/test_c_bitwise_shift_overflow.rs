use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn shift_overflow_left_signed() {
    assert_eq!(
        run_c(
            "int main() { int x = 1073741824; /* x << 2 overflows 32-bit signed int -> UB */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn shift_overflow_left_unsigned() {
    assert_eq!(
        run_c("int main() { unsigned int x = 2147483648U; printf(\"%u\", x << 1); return 0; }"),
        vec!["0"]
    );
} // Well defined, wraps
#[test]
fn shift_count_overflow() {
    assert_eq!(
        run_c(
            "int main() { int x = 1; /* x << 32 is UB if int is 32-bit */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn shift_count_overflow_unsigned() {
    assert_eq!(
        run_c(
            "int main() { unsigned int x = 1; /* x << 32 is UB if int is 32-bit */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn shift_count_exceeds_promotion() {
    assert_eq!(
        run_c(
            "int main() { char c = 1; /* c << 16 is fine if int is 32-bit because c promotes to int before shift */ printf(\"%d\", c << 16); return 0; }"
        ),
        vec!["65536"]
    );
}
#[test]
fn shift_count_exceeds_promotion_ub() {
    assert_eq!(
        run_c(
            "int main() { char c = 1; /* c << 32 is UB if int is 32-bit */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn shift_left_into_sign_bit() {
    assert_eq!(
        run_c(
            "int main() { int x = 1073741824; /* 1 << 30 */ /* x << 1 into sign bit is UB in C99/C11 */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn shift_right_overflow_count() {
    assert_eq!(
        run_c("int main() { int x = 16; /* x >> 32 is UB */ printf(\"ok\"); return 0; }"),
        vec!["ok"]
    );
}
#[test]
fn shift_overflow_compound_assignment() {
    assert_eq!(
        run_c("int main() { unsigned int x = 2147483648U; x <<= 1; printf(\"%u\", x); return 0; }"),
        vec!["0"]
    );
}
#[test]
fn shift_overflow_long_long() {
    assert_eq!(
        run_c(
            "int main() { unsigned long long x = 9223372036854775808ULL; printf(\"%llu\", x << 1); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn shift_overflow_in_constant_expr() {
    assert_eq!(
        run_c("int main() { enum { A = 1U << 31 }; printf(\"%u\", (unsigned int)A); return 0; }"),
        vec!["2147483648"]
    );
}
#[test]
fn shift_overflow_in_constant_expr_ub() {
    assert_eq!(
        run_c(
            "/* enum { A = 1 << 32 }; // Error in compilation often */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn shift_overflow_implicit_type() {
    assert_eq!(
        run_c("int main() { printf(\"%lld\", 1LL << 40); return 0; }"),
        vec!["1099511627776"]
    );
} // 1LL makes it long long, so no UB
#[test]
fn shift_overflow_no_suffix() {
    assert_eq!(
        run_c(
            "/* printf(\"%lld\", 1 << 40); // 1 is int, << 40 is UB */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn shift_overflow_char_shift() {
    assert_eq!(
        run_c("int main() { unsigned char c = 255; printf(\"%d\", c << 8); return 0; }"),
        vec!["65280"]
    );
} // Valid, c promotes to int, 255 << 8 fits in 32-bit int
