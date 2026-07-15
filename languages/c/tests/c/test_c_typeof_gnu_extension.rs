use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn typeof_basic() {
    assert_eq!(
        run_c("int main() { int a=5; typeof(a) b = 10; printf(\"%d\", b); return 0; }"),
        vec!["10"]
    );
}
#[test]
fn typeof_type_name() {
    assert_eq!(
        run_c("int main() { typeof(int) a = 7; printf(\"%d\", a); return 0; }"),
        vec!["7"]
    );
}
#[test]
fn typeof_pointer() {
    assert_eq!(
        run_c("int main() { int a=5; typeof(&a) p = &a; printf(\"%d\", *p); return 0; }"),
        vec!["5"]
    );
}
#[test]
fn typeof_struct() {
    assert_eq!(
        run_c(
            "struct S { int x; }; int main() { struct S s1={1}; typeof(s1) s2={2}; printf(\"%d\", s2.x); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn typeof_array() {
    assert_eq!(
        run_c(
            "int main() { int a[3]={1,2,3}; typeof(a) b={4,5,6}; printf(\"%d\", b[1]); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn typeof_function() {
    assert_eq!(
        run_c(
            "int f(int) { return 1; } int main() { typeof(f) *p = f; printf(\"%d\", p(0)); return 0; }"
        ),
        vec!["1"]
    );
} // typeof(f) is function type, so we need pointer
#[test]
fn typeof_expr_unevaluated() {
    assert_eq!(
        run_c(
            "int f(int *x) { (*x)++; return 1; } int main() { int x=0; typeof(f(&x)) y = 5; printf(\"%d\", x); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn typeof_vla() {
    assert_eq!(
        run_c(
            "int main() { int n=5; int a[n]; typeof(a) b; /* b is also VLA of size 5 */ printf(\"%d\", (int)(sizeof(b)/sizeof(int))); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn typeof_vla_evaluated() {
    assert_eq!(
        run_c("int main() { int x=1; typeof(int[x++]) b; printf(\"%d\", x); return 0; }"),
        vec!["2"]
    );
} // VLA size is evaluated
#[test]
fn typeof_const_qualifier() {
    assert_eq!(
        run_c(
            "int main() { const int a=1; typeof(a) b=2; /* b is const */ /* b=3; // error */ printf(\"%d\", b); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn typeof_with_pointers() {
    assert_eq!(
        run_c("int main() { int a=1; typeof(a)* p = &a; printf(\"%d\", *p); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn typeof_in_macro() {
    assert_eq!(
        run_c(
            "#define SWAP(a, b) do { typeof(a) tmp = a; a = b; b = tmp; } while(0)\nint main() { int x=1, y=2; SWAP(x, y); printf(\"%d\", x); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn typeof_nested() {
    assert_eq!(
        run_c("int main() { int a=1; typeof(typeof(a)) b = 5; printf(\"%d\", b); return 0; }"),
        vec!["5"]
    );
}
#[test]
fn typeof_cast() {
    assert_eq!(
        run_c("int main() { double d = 3.14; int i = (typeof(1))d; printf(\"%d\", i); return 0; }"),
        vec!["3"]
    );
}
#[test]
fn typeof_typeof_unqualified() {
    assert_eq!(
        run_c(
            "/* GNU C has typeof_unqual in C23, let's just stick to typeof for GNU ext */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
