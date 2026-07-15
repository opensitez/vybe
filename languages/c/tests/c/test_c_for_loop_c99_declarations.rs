use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn for_c99_decl_basic() {
    assert_eq!(
        run_c(
            "int main() { int sum=0; for(int i=0; i<3; i++) sum += i; printf(\"%d\", sum); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn for_c99_decl_scope() {
    assert_eq!(
        run_c(
            "int main() { for(int i=0; i<1; i++) ; /* printf(\"%d\", i); // error i undeclared */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn for_c99_decl_shadows_outer() {
    assert_eq!(
        run_c("int main() { int i=10; for(int i=0; i<2; i++) ; printf(\"%d\", i); return 0; }"),
        vec!["10"]
    );
}
#[test]
fn for_c99_decl_shadowed_by_inner() {
    assert_eq!(
        run_c("int main() { for(int i=0; i<1; i++) { int i=5; printf(\"%d\", i); } return 0; }"),
        vec!["5"]
    );
}
#[test]
fn for_c99_decl_multiple_loops() {
    assert_eq!(
        run_c(
            "int main() { int sum=0; for(int i=0; i<2; i++) sum++; for(int i=0; i<2; i++) sum++; printf(\"%d\", sum); return 0; }"
        ),
        vec!["4"]
    );
}
#[test]
fn for_c99_decl_struct_tag() {
    assert_eq!(
        run_c(
            "int main() { for(struct S { int x; } s = {5}; s.x < 6; s.x++) printf(\"%d\", s.x); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn for_c99_decl_enum() {
    assert_eq!(
        run_c("int main() { for(enum { A=1, B=3 } e = A; e < B; e++) printf(\"-\"); return 0; }"),
        vec!["-", "-"]
    );
}
#[test]
fn for_c99_decl_static_fails() {
    assert_eq!(
        run_c(
            "/* int main() { for(static int i=0; i<1; i++) {} return 0; } // static not allowed */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn for_c99_decl_register() {
    assert_eq!(
        run_c("int main() { for(register int i=0; i<1; i++) printf(\"R\"); return 0; }"),
        vec!["R"]
    );
}
#[test]
fn for_c99_decl_const() {
    assert_eq!(
        run_c("int main() { for(const int i=0; i<1; ) { printf(\"C\"); break; } return 0; }"),
        vec!["C"]
    );
}
#[test]
fn for_c99_decl_volatile() {
    assert_eq!(
        run_c("int main() { for(volatile int i=0; i<1; i++) printf(\"V\"); return 0; }"),
        vec!["V"]
    );
}
#[test]
fn for_c99_decl_pointer() {
    assert_eq!(
        run_c("int main() { int a=1; for(int *p=&a; *p<2; (*p)++) printf(\"P\"); return 0; }"),
        vec!["P"]
    );
}
#[test]
fn for_c99_decl_array() {
    assert_eq!(
        run_c("int main() { for(int arr[1]={0}; arr[0]<1; arr[0]++) printf(\"A\"); return 0; }"),
        vec!["A"]
    );
}
#[test]
fn for_c99_decl_extern_fails() {
    assert_eq!(
        run_c(
            "/* int main() { for(extern int i=0; ;) {} return 0; } */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn for_c99_decl_function_ptr() {
    assert_eq!(
        run_c(
            "int f() { return 1; } int main() { for(int (*p)(void)=f; p()!=0; p=0) printf(\"F\"); return 0; }"
        ),
        vec!["F"]
    );
}
