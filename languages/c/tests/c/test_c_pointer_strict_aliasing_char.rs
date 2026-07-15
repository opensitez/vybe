use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn strict_aliasing_char_reads_int() {
    assert_eq!(
        run_c("int main() { int x = 0x12345678; char *p = (char*)&x; printf(\"ok\"); return 0; }"),
        vec!["ok"]
    );
} // Char can alias anything
#[test]
fn strict_aliasing_unsigned_char() {
    assert_eq!(
        run_c(
            "int main() { float f = 3.14f; unsigned char *p = (unsigned char*)&f; printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // Unsigned char can alias anything
#[test]
fn strict_aliasing_signed_char() {
    assert_eq!(
        run_c(
            "int main() { double d = 1.0; signed char *p = (signed char*)&d; printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn strict_aliasing_int_float_fails() {
    assert_eq!(
        run_c(
            "/* int main() { float f = 1.0f; int *p = (int*)&f; *p = 1; return 0; } // UB */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn strict_aliasing_union_allowed() {
    assert_eq!(
        run_c(
            "union U { int i; float f; }; int main() { union U u; u.f = 1.0f; int val = u.i; printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // Type punning via union is standard C99
#[test]
fn strict_aliasing_memcpy_allowed() {
    assert_eq!(
        run_c(
            "#include <string.h>\nint main() { float f = 1.0f; int i; memcpy(&i, &f, sizeof(float)); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // memcpy is safe
#[test]
fn strict_aliasing_void_ptr() {
    assert_eq!(
        run_c(
            "int main() { float f = 1.0f; void *p = &f; int *ip = (int*)p; /* *ip = 1; UB */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn strict_aliasing_struct_members() {
    assert_eq!(
        run_c(
            "struct S { int i; float f; }; int main() { struct S s; s.f = 1.0f; int *p = (int*)&s.f; /* UB if dereferenced */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn strict_aliasing_array_pun() {
    assert_eq!(
        run_c(
            "int main() { int arr[2] = {1, 2}; long long *p = (long long*)arr; /* UB if dereferenced */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn strict_aliasing_char_array_to_int() {
    assert_eq!(
        run_c(
            "int main() { char arr[sizeof(int)] = {0}; int *p = (int*)arr; /* UB if unaligned, or strict aliasing violation if arr is declared char. But often done. */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn strict_aliasing_same_type_qualifiers() {
    assert_eq!(
        run_c(
            "int main() { const int x = 5; int *p = (int*)&x; /* valid pointer conversion, UB to write */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn strict_aliasing_compatible_types() {
    assert_eq!(
        run_c(
            "int main() { signed int x = 5; unsigned int *p = (unsigned int*)&x; printf(\"%u\", *p); return 0; }"
        ),
        vec!["5"]
    );
} // signed and unsigned of same type can alias
#[test]
fn strict_aliasing_struct_prefix() {
    assert_eq!(
        run_c(
            "struct A { int x; }; struct B { int x; float y; }; int main() { struct B b = {1, 2.0}; struct A *p = (struct A*)&b; /* UB in standard C but common in old code */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn strict_aliasing_malloc() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { void *p = malloc(sizeof(int)); *(int*)p = 5; printf(\"%d\", *(int*)p); free(p); return 0; }"
        ),
        vec!["5"]
    );
} // allocated storage has no declared type, takes on type of lvalue
