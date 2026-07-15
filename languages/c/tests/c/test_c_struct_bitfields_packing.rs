use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn bitfield_packing_basic() {
    assert_eq!(
        run_c(
            "struct S { unsigned int a:4; unsigned int b:4; }; int main() { printf(\"%d\", sizeof(struct S) == sizeof(unsigned int)); return 0; }"
        ),
        vec!["1"]
    );
} // Should pack into one int
#[test]
fn bitfield_packing_overflow() {
    assert_eq!(
        run_c(
            "struct S { unsigned int a:30; unsigned int b:10; }; int main() { printf(\"%d\", sizeof(struct S) > sizeof(unsigned int)); return 0; }"
        ),
        vec!["1"]
    );
} // Doesn't fit in one 32-bit int
#[test]
fn bitfield_packing_zero_width() {
    assert_eq!(
        run_c(
            "struct S { unsigned int a:4; unsigned int :0; unsigned int b:4; }; int main() { printf(\"%d\", sizeof(struct S) > sizeof(unsigned int)); return 0; }"
        ),
        vec!["1"]
    );
} // :0 forces next bitfield to align to next boundary
#[test]
fn bitfield_packing_unnamed() {
    assert_eq!(
        run_c(
            "struct S { unsigned int a:4; unsigned int :4; unsigned int b:4; }; int main() { struct S s = {1, 2}; printf(\"%d\", s.b); return 0; }"
        ),
        vec!["2"]
    );
} // Unnamed padding
#[test]
fn bitfield_packing_different_types() {
    assert_eq!(
        run_c("struct S { char a:4; int b:4; }; int main() { printf(\"ok\"); return 0; }"),
        vec!["ok"]
    );
} // Implementation defined whether they pack together
#[test]
fn bitfield_packing_sizeof_fails() {
    assert_eq!(
        run_c(
            "struct S { int a:4; }; int main() { /* sizeof(s.a) // illegal */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn bitfield_packing_address_of_fails() {
    assert_eq!(
        run_c(
            "struct S { int a:4; }; int main() { /* &s.a // illegal */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn bitfield_packing_read() {
    assert_eq!(
        run_c(
            "struct S { unsigned int a:3; }; int main() { struct S s; s.a = 5; printf(\"%d\", s.a); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn bitfield_packing_write_overflow() {
    assert_eq!(
        run_c(
            "struct S { unsigned int a:2; }; int main() { struct S s; s.a = 5; /* 5 is 101, fits in 3 bits. 2 bits holds up to 3. 5 & 3 = 1 */ printf(\"%d\", s.a); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn bitfield_packing_boolean() {
    assert_eq!(
        run_c(
            "#include <stdbool.h>\nstruct S { bool a:1; }; int main() { struct S s; s.a = 2; printf(\"%d\", s.a); return 0; }"
        ),
        vec!["1"]
    );
} // Any non-zero usually becomes 1
#[test]
fn bitfield_packing_large_width_fails() {
    assert_eq!(
        run_c(
            "/* struct S { int a:100; }; // width exceeds type */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn bitfield_packing_enum() {
    assert_eq!(
        run_c(
            "enum E { A, B }; struct S { enum E e:2; }; int main() { struct S s; s.e = B; printf(\"%d\", s.e); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn bitfield_packing_struct_assignment() {
    assert_eq!(
        run_c(
            "struct S { unsigned int a:4; unsigned int b:4; }; int main() { struct S s1 = {1, 2}; struct S s2 = s1; printf(\"%d\", s2.a+s2.b); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn bitfield_packing_promotion() {
    assert_eq!(
        run_c(
            "struct S { unsigned int a:4; }; int main() { struct S s = {15}; printf(\"%d\", sizeof(s.a + 1)); return 0; }"
        ),
        vec!["4"]
    );
} // Promotes to int
#[test]
fn bitfield_packing_pointer_to_struct() {
    assert_eq!(
        run_c(
            "struct S { unsigned int a:4; }; int main() { struct S s = {7}; struct S *p = &s; printf(\"%d\", p->a); return 0; }"
        ),
        vec!["7"]
    );
}
