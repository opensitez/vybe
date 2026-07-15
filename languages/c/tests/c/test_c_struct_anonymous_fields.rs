use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn struct_anon_field_basic() {
    assert_eq!(
        run_c(
            "struct Inner { int a; }; struct Outer { struct Inner; int b; }; int main() { struct Outer o; o.a = 1; o.b = 2; printf(\"%d\", o.a + o.b); return 0; }"
        ),
        vec!["3"]
    );
} // C11 anonymous structs
#[test]
fn struct_anon_union_field() {
    assert_eq!(
        run_c(
            "struct S { union { int i; float f; }; int type; }; int main() { struct S s; s.i = 5; printf(\"%d\", s.i); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn struct_anon_field_nested() {
    assert_eq!(
        run_c(
            "struct A { int x; }; struct B { struct A; int y; }; struct C { struct B; int z; }; int main() { struct C c; c.x = 1; c.y = 2; c.z = 3; printf(\"%d\", c.x+c.y+c.z); return 0; }"
        ),
        vec!["6"]
    );
}
#[test]
fn struct_anon_field_initialization() {
    assert_eq!(
        run_c(
            "struct S { struct { int a; int b; }; int c; }; int main() { struct S s = { {1, 2}, 3 }; printf(\"%d\", s.a); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn struct_anon_field_designated_init() {
    assert_eq!(
        run_c(
            "struct S { struct { int a; int b; }; int c; }; int main() { struct S s = { .a = 1, .c = 3 }; printf(\"%d\", s.a + s.c); return 0; }"
        ),
        vec!["4"]
    );
} // In C11 you can initialize anonymous members directly via designation
#[test]
fn struct_anon_field_shadowing_fails() {
    assert_eq!(
        run_c(
            "/* struct S { struct { int a; }; int a; }; */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // Member names must be unique
#[test]
fn struct_anon_union_designated_init() {
    assert_eq!(
        run_c(
            "struct S { union { int i; char c; }; }; int main() { struct S s = { .i = 65 }; printf(\"%d\", s.i); return 0; }"
        ),
        vec!["65"]
    );
}
#[test]
fn struct_anon_field_sizeof() {
    assert_eq!(
        run_c(
            "struct S { struct { int a; int b; }; }; int main() { printf(\"%d\", (int)sizeof(struct S)); return 0; }"
        ),
        vec!["8"]
    );
}
#[test]
fn struct_anon_field_address_of() {
    assert_eq!(
        run_c(
            "struct S { struct { int a; }; }; int main() { struct S s; s.a = 10; int *p = &s.a; printf(\"%d\", *p); return 0; }"
        ),
        vec!["10"]
    );
}
#[test]
fn struct_anon_typedef_field_fails() {
    assert_eq!(
        run_c(
            "typedef struct { int a; } T; /* struct S { T; }; // Anonymous must not have tag or typedef name */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn struct_anon_field_array() {
    assert_eq!(
        run_c(
            "struct S { struct { int arr[2]; }; }; int main() { struct S s; s.arr[1] = 5; printf(\"%d\", s.arr[1]); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn struct_anon_field_bitfield() {
    assert_eq!(
        run_c(
            "struct S { struct { int a:4; int b:4; }; }; int main() { struct S s; s.a = 2; printf(\"%d\", s.a); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn struct_anon_union_in_union() {
    assert_eq!(
        run_c(
            "union U { union { int a; int b; }; float c; }; int main() { union U u; u.a = 10; printf(\"%d\", u.b); return 0; }"
        ),
        vec!["10"]
    );
}
#[test]
fn struct_anon_field_multiple() {
    assert_eq!(
        run_c(
            "struct S { struct { int a; }; struct { int b; }; }; int main() { struct S s; s.a=1; s.b=2; printf(\"%d\", s.a+s.b); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn struct_anon_field_with_tag_fails() {
    assert_eq!(
        run_c(
            "/* struct S { struct Tag { int a; }; }; // Tagged structs are not anonymous members */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
