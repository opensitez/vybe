use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn compound_literal_basic_struct() {
    assert_eq!(
        run_c(
            "struct S { int a, b; }; int main() { struct S s = (struct S){1, 2}; printf(\"%d\", s.a + s.b); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn compound_literal_basic_array() {
    assert_eq!(
        run_c("int main() { int *p = (int[]){1, 2, 3}; printf(\"%d\", p[1]); return 0; }"),
        vec!["2"]
    );
}
#[test]
fn compound_literal_scalar() {
    assert_eq!(
        run_c("int main() { int *p = &(int){42}; printf(\"%d\", *p); return 0; }"),
        vec!["42"]
    );
}
#[test]
fn compound_literal_as_argument() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int f(struct S s) { return s.a; } int main() { printf(\"%d\", f((struct S){5})); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn compound_literal_pointer_return() {
    assert_eq!(
        run_c(
            "int* f() { return (int[]){10, 20}; } int main() { /* UB if returned from function, testing parsing/scope conceptually */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn compound_literal_file_scope() {
    assert_eq!(
        run_c(
            "struct S { int a; }; struct S *p = &(struct S){99}; int main() { printf(\"%d\", p->a); return 0; }"
        ),
        vec!["99"]
    );
}
#[test]
fn compound_literal_const() {
    assert_eq!(
        run_c("int main() { const int *p = &(const int){77}; printf(\"%d\", *p); return 0; }"),
        vec!["77"]
    );
}
#[test]
fn compound_literal_designated() {
    assert_eq!(
        run_c(
            "struct S { int a, b; }; int main() { struct S s = (struct S){.b = 10, .a = 5}; printf(\"%d\", s.b - s.a); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn compound_literal_array_designated() {
    assert_eq!(
        run_c("int main() { int *p = (int[]){[2] = 5, [0] = 1}; printf(\"%d\", p[2]); return 0; }"),
        vec!["5"]
    );
}
#[test]
fn compound_literal_modification() {
    assert_eq!(
        run_c("int main() { int *p = (int[]){1, 2}; p[0] = 10; printf(\"%d\", p[0]); return 0; }"),
        vec!["10"]
    );
} // They are lvalues
#[test]
fn compound_literal_sizeof() {
    assert_eq!(
        run_c("int main() { printf(\"%d\", (int)sizeof((int[]){1,2,3,4})); return 0; }"),
        vec!["16"]
    );
} // Assuming 4-byte int
#[test]
fn compound_literal_address_of() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { struct S *p = &(struct S){12}; printf(\"%d\", p->a); return 0; }"
        ),
        vec!["12"]
    );
}
#[test]
fn compound_literal_nested_in_expr() {
    assert_eq!(
        run_c("int main() { int x = ((int[]){1, 2, 3})[2]; printf(\"%d\", x); return 0; }"),
        vec!["3"]
    );
}
#[test]
fn compound_literal_loop_scope() {
    assert_eq!(
        run_c(
            "int main() { int *p; for (int i=0; i<1; i++) { p = &(int){100}; } printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // Testing parsing
#[test]
fn compound_literal_union() {
    assert_eq!(
        run_c(
            "union U { int i; float f; }; int main() { union U *u = &(union U){.i = 55}; printf(\"%d\", u->i); return 0; }"
        ),
        vec!["55"]
    );
}
