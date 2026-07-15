use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn compound_literal_nested_in_struct() {
    assert_eq!(
        run_c(
            "struct Inner { int x; }; struct Outer { struct Inner i; }; int main() { struct Outer o = { (struct Inner){5} }; printf(\"%d\", o.i.x); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn compound_literal_nested_in_array() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { struct S arr[2] = { (struct S){1}, (struct S){2} }; printf(\"%d\", arr[1].a); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn compound_literal_pointer_in_struct() {
    assert_eq!(
        run_c(
            "struct S { int *p; }; int main() { struct S s = { (int[]){10, 20} }; printf(\"%d\", s.p[1]); return 0; }"
        ),
        vec!["20"]
    );
}
#[test]
fn compound_literal_double_nested() {
    assert_eq!(
        run_c(
            "struct A { int a; }; struct B { struct A *p; }; struct C { struct B b; }; int main() { struct C c = { (struct B){ &(struct A){42} } }; printf(\"%d\", c.b.p->a); return 0; }"
        ),
        vec!["42"]
    );
}
#[test]
fn compound_literal_in_ternary() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { int cond = 1; struct S s = cond ? (struct S){1} : (struct S){2}; printf(\"%d\", s.a); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn compound_literal_in_ternary_false() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { int cond = 0; struct S s = cond ? (struct S){1} : (struct S){2}; printf(\"%d\", s.a); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn compound_literal_passed_to_macro() {
    assert_eq!(
        run_c(
            "#define GET_A(s) ((s).a)\nstruct S { int a; }; int main() { printf(\"%d\", GET_A((struct S){99})); return 0; }"
        ),
        vec!["99"]
    );
}
#[test]
fn compound_literal_array_of_pointers() {
    assert_eq!(
        run_c(
            "int main() { int *arr[] = { (int[]){1}, (int[]){2} }; printf(\"%d\", arr[0][0] + arr[1][0]); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn compound_literal_struct_of_arrays() {
    assert_eq!(
        run_c(
            "struct S { int arr[2]; }; int main() { struct S *s = &(struct S){ {5, 6} }; printf(\"%d\", s->arr[0]); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn compound_literal_in_switch() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { switch(((struct S){2}).a) { case 2: printf(\"2\"); break; } return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn compound_literal_in_for_init() {
    assert_eq!(
        run_c(
            "int main() { for (int *p = (int[]){0}; *p < 2; (*p)++) { printf(\"%d\", *p); } return 0; }"
        ),
        vec!["0", "1"]
    );
}
#[test]
fn compound_literal_pointer_to_function() {
    assert_eq!(
        run_c(
            "int f() { return 1; } int main() { int (*p)(void) = ((int (*[])(void)){f})[0]; printf(\"%d\", p()); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn compound_literal_in_sizeof() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { printf(\"%d\", (int)sizeof((struct S){1})); return 0; }"
        ),
        vec!["4"]
    );
} // Assuming 4 byte int
#[test]
fn compound_literal_compound_assignment() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { struct S s = {1}; s = (struct S){s.a + 2}; printf(\"%d\", s.a); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn compound_literal_comma_operator() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { struct S s = (1, (struct S){5}); printf(\"%d\", s.a); return 0; }"
        ),
        vec!["5"]
    );
}
