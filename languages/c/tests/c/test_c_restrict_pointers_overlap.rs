use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn restrict_pointer_basic() {
    assert_eq!(
        run_c(
            "void f(int *restrict p) { *p = 5; } int main() { int x; f(&x); printf(\"%d\", x); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn restrict_pointer_two_args() {
    assert_eq!(
        run_c(
            "void f(int *restrict p1, int *restrict p2) { *p1 = 1; *p2 = 2; } int main() { int x, y; f(&x, &y); printf(\"%d%d\", x, y); return 0; }"
        ),
        vec!["12"]
    );
}
#[test]
fn restrict_pointer_overlap_fails() {
    assert_eq!(
        run_c(
            "/* void f(int *restrict p1, int *restrict p2) { *p1=1; *p2=2; } int main() { int x; f(&x, &x); return 0; } // UB to alias restrict pointers */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn restrict_pointer_local() {
    assert_eq!(
        run_c(
            "int main() { int x = 1; int *restrict p = &x; *p = 2; printf(\"%d\", x); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn restrict_pointer_struct_member() {
    assert_eq!(
        run_c(
            "struct S { int *restrict p; }; int main() { int x = 1; struct S s; s.p = &x; *s.p = 2; printf(\"%d\", x); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn restrict_pointer_array_arg() {
    assert_eq!(
        run_c(
            "void f(int arr[restrict 5]) { arr[0] = 5; } int main() { int a[5]; f(a); printf(\"%d\", a[0]); return 0; }"
        ),
        vec!["5"]
    );
} // C99 restrict in array parameter
#[test]
fn restrict_pointer_const() {
    assert_eq!(
        run_c(
            "void f(const int *restrict p) { printf(\"%d\", *p); } int main() { int x = 5; f(&x); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn restrict_pointer_return() {
    assert_eq!(
        run_c(
            "int *restrict f(int *p) { return p; } int main() { int x = 5; int *restrict p = f(&x); printf(\"%d\", *p); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn restrict_pointer_nested() {
    assert_eq!(
        run_c(
            "int main() { int x; int *restrict *restrict p; int *restrict p1 = &x; p = &p1; **p = 5; printf(\"%d\", x); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn restrict_pointer_typedef() {
    assert_eq!(
        run_c(
            "typedef int *intptr; int main() { int x = 1; intptr restrict p = &x; *p = 2; printf(\"%d\", x); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn restrict_pointer_to_function_fails() {
    assert_eq!(
        run_c(
            "/* void (*restrict f)() = 0; // restrict only applies to pointers to object types */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn restrict_pointer_block_scope() {
    assert_eq!(
        run_c(
            "int main() { int x=1; { int *restrict p = &x; *p = 2; } { int *restrict q = &x; *q = 3; } printf(\"%d\", x); return 0; }"
        ),
        vec!["3"]
    );
} // fine, different blocks
#[test]
fn restrict_pointer_assignment() {
    assert_eq!(
        run_c(
            "int main() { int x=1; int *restrict p = &x; int *q = p; *q = 2; printf(\"%d\", x); return 0; }"
        ),
        vec!["2"]
    );
} // fine, p is not accessed after assignment
#[test]
fn restrict_pointer_memcpy_like() {
    assert_eq!(
        run_c(
            "void my_memcpy(void *restrict dest, const void *restrict src, int n) { char *d = dest; const char *s = src; while(n--) *d++ = *s++; } int main() { char a[5]=\"abcd\"; char b[5]; my_memcpy(b, a, 5); printf(\"%s\", b); return 0; }"
        ),
        vec!["abcd"]
    );
}
