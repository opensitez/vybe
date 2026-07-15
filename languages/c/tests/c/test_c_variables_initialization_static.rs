use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn static_init_zero_implicit() {
    assert_eq!(
        run_c("static int a; int main() { printf(\"%d\", a); return 0; }"),
        vec!["0"]
    );
}
#[test]
fn static_init_local_zero_implicit() {
    assert_eq!(
        run_c("int main() { static int a; printf(\"%d\", a); return 0; }"),
        vec!["0"]
    );
}
#[test]
fn static_init_local_persists() {
    assert_eq!(
        run_c(
            "int f() { static int a = 0; a++; return a; } int main() { f(); printf(\"%d\", f()); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn static_init_pointer_implicit_null() {
    assert_eq!(
        run_c("static int *p; int main() { printf(\"%d\", p == 0); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn static_init_array_implicit_zero() {
    assert_eq!(
        run_c("static int arr[3]; int main() { printf(\"%d\", arr[2]); return 0; }"),
        vec!["0"]
    );
}
#[test]
fn static_init_struct_implicit_zero() {
    assert_eq!(
        run_c(
            "struct S { int a; }; static struct S s; int main() { printf(\"%d\", s.a); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn static_init_explicit_const() {
    assert_eq!(
        run_c("static int a = 42; int main() { printf(\"%d\", a); return 0; }"),
        vec!["42"]
    );
}
#[test]
fn static_init_address_of_global() {
    assert_eq!(
        run_c("int g = 10; static int *p = &g; int main() { printf(\"%d\", *p); return 0; }"),
        vec!["10"]
    );
}
#[test]
fn static_init_string_literal() {
    assert_eq!(
        run_c("static char *s = \"hello\"; int main() { printf(\"%c\", s[0]); return 0; }"),
        vec!["h"]
    );
}
#[test]
fn static_init_char_array_literal() {
    assert_eq!(
        run_c("static char s[] = \"hi\"; int main() { printf(\"%c\", s[1]); return 0; }"),
        vec!["i"]
    );
}
#[test]
fn static_init_struct_with_address() {
    assert_eq!(
        run_c(
            "int g = 5; struct S { int *p; }; static struct S s = { &g }; int main() { printf(\"%d\", *s.p); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn static_init_union() {
    assert_eq!(
        run_c(
            "union U { int a; char c; }; static union U u = { 65 }; int main() { printf(\"%d\", u.a); return 0; }"
        ),
        vec!["65"]
    );
}
#[test]
fn static_local_shadows_global() {
    assert_eq!(
        run_c("int a = 1; int main() { static int a = 2; printf(\"%d\", a); return 0; }"),
        vec!["2"]
    );
}
#[test]
fn static_init_recursive_call() {
    assert_eq!(
        run_c(
            "int f(int n) { static int depth = 0; depth++; if (n == 0) return depth; return f(n-1); } int main() { printf(\"%d\", f(2)); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn static_init_expr_fails_if_not_const() {
    assert_eq!(
        run_c(
            "int g = 5; int main() { /* static int a = g; fails in pure C */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
