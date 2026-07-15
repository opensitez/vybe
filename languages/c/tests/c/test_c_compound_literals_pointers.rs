use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn compound_literal_pointer_basic() {
    assert_eq!(
        run_c("int main() { int *p = &(int){5}; printf(\"%d\", *p); return 0; }"),
        vec!["5"]
    );
}
#[test]
fn compound_literal_pointer_struct() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { struct S *p = &(struct S){42}; printf(\"%d\", p->a); return 0; }"
        ),
        vec!["42"]
    );
}
#[test]
fn compound_literal_pointer_array() {
    assert_eq!(
        run_c("int main() { int *p = (int[]){1, 2, 3}; printf(\"%d\", p[2]); return 0; }"),
        vec!["3"]
    );
} // Array decays to pointer naturally
#[test]
fn compound_literal_pointer_lifetime_block() {
    assert_eq!(
        run_c(
            "int main() { int *p; { p = &(int){7}; } /* *p might be invalid here, but often works in practice. let's just test compiling */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn compound_literal_pointer_lifetime_file() {
    assert_eq!(
        run_c("int *p = &(int){99}; int main() { printf(\"%d\", *p); return 0; }"),
        vec!["99"]
    );
} // Has static storage duration
#[test]
fn compound_literal_pointer_const() {
    assert_eq!(
        run_c("int main() { const int *p = &(const int){10}; printf(\"%d\", *p); return 0; }"),
        vec!["10"]
    );
} // Can be put in ROM
#[test]
fn compound_literal_pointer_modification() {
    assert_eq!(
        run_c("int main() { int *p = &(int){10}; *p = 20; printf(\"%d\", *p); return 0; }"),
        vec!["20"]
    );
} // Modifiable if not const
#[test]
fn compound_literal_pointer_to_pointer() {
    assert_eq!(
        run_c("int main() { int **p = &(int*){ &(int){5} }; printf(\"%d\", **p); return 0; }"),
        vec!["5"]
    );
}
#[test]
fn compound_literal_pointer_function_arg() {
    assert_eq!(
        run_c("void f(int *p) { printf(\"%d\", *p); } int main() { f(&(int){15}); return 0; }"),
        vec!["15"]
    );
}
#[test]
fn compound_literal_pointer_function_return() {
    assert_eq!(
        run_c(
            "int *f() { return &(int){5}; /* Returns pointer to local, UB, but compiles */ } int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn compound_literal_pointer_array_of_pointers() {
    assert_eq!(
        run_c(
            "int main() { int *arr[] = { &(int){1}, &(int){2} }; printf(\"%d\", *arr[1]); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn compound_literal_pointer_in_struct() {
    assert_eq!(
        run_c(
            "struct S { int *p; }; int main() { struct S s = { &(int){77} }; printf(\"%d\", *s.p); return 0; }"
        ),
        vec!["77"]
    );
}
#[test]
fn compound_literal_pointer_ternary() {
    assert_eq!(
        run_c("int main() { int *p = 1 ? &(int){1} : &(int){2}; printf(\"%d\", *p); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn compound_literal_pointer_comma() {
    assert_eq!(
        run_c("int main() { int *p = (1, &(int){5}); printf(\"%d\", *p); return 0; }"),
        vec!["5"]
    );
}
#[test]
fn compound_literal_pointer_cast() {
    assert_eq!(
        run_c("int main() { void *p = &(int){88}; printf(\"%d\", *(int*)p); return 0; }"),
        vec!["88"]
    );
}
