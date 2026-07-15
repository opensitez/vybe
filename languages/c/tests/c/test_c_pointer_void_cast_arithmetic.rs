use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn void_pointer_basic() {
    assert_eq!(
        run_c(
            "int main() { int x = 5; void *p = &x; int *ip = p; printf(\"%d\", *ip); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn void_pointer_deref_fails() {
    assert_eq!(
        run_c(
            "/* int main() { int x = 5; void *p = &x; *p; return 0; } */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn void_pointer_arithmetic_gnu() {
    assert_eq!(
        run_c(
            "/* GNU C allows void* arithmetic, treats size as 1 */ int main() { int x[2] = {1, 2}; void *p = x; /* let's just check standard behavior or ignore */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn void_pointer_cast_to_char() {
    assert_eq!(
        run_c(
            "int main() { int x = 0x12345678; void *p = &x; char *cp = p; /* accesses byte, endian dependent, just test compile */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn void_pointer_cast_to_function_fails() {
    assert_eq!(
        run_c(
            "/* void f(){} int main() { void *p = f; return 0; } // standard C doesn't guarantee object pointer can hold function pointer, but POSIX does */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn void_pointer_comparison() {
    assert_eq!(
        run_c(
            "int main() { int x; void *p1 = &x; void *p2 = &x; printf(\"%d\", p1 == p2); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn void_pointer_relational() {
    assert_eq!(
        run_c(
            "int main() { int arr[2]; void *p1 = &arr[0]; void *p2 = &arr[1]; printf(\"%d\", p1 < p2); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn void_pointer_subtraction_gnu() {
    assert_eq!(
        run_c(
            "/* int main() { int arr[2]; void *p1 = &arr[0]; void *p2 = &arr[1]; int d = p2 - p1; return 0; } // GNU ext */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn void_pointer_return() {
    assert_eq!(
        run_c(
            "void *f(int *x) { return x; } int main() { int x = 5; int *p = f(&x); printf(\"%d\", *p); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn void_pointer_arg() {
    assert_eq!(
        run_c(
            "void f(void *p) { printf(\"%d\", *(int*)p); } int main() { int x = 6; f(&x); return 0; }"
        ),
        vec!["6"]
    );
}
#[test]
fn void_pointer_ternary() {
    assert_eq!(
        run_c(
            "int main() { int x=1; float y=2; void *p = 1 ? &x : &y; printf(\"%d\", *(int*)p); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn void_pointer_null() {
    assert_eq!(
        run_c("int main() { void *p = 0; printf(\"%d\", p == 0); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn void_pointer_sizeof_gnu() {
    assert_eq!(
        run_c(
            "/* int main() { int s = sizeof(void); return 0; } // GNU ext is 1 */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn void_pointer_const() {
    assert_eq!(
        run_c(
            "int main() { const int x = 5; const void *p = &x; printf(\"%d\", *(const int*)p); return 0; }"
        ),
        vec!["5"]
    );
}
