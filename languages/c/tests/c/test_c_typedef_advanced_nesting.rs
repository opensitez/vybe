use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn typedef_basic() {
    assert_eq!(
        run_c("typedef int myint; int main() { myint a = 1; printf(\"%d\", a); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn typedef_pointer() {
    assert_eq!(
        run_c(
            "typedef int* intptr; int main() { int a = 2; intptr p = &a; printf(\"%d\", *p); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn typedef_array() {
    assert_eq!(
        run_c(
            "typedef int intarr[3]; int main() { intarr a = {1, 2, 3}; printf(\"%d\", a[1]); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn typedef_struct() {
    assert_eq!(
        run_c(
            "typedef struct { int a; } S; int main() { S s = {4}; printf(\"%d\", s.a); return 0; }"
        ),
        vec!["4"]
    );
}
#[test]
fn typedef_struct_with_tag() {
    assert_eq!(
        run_c(
            "typedef struct Tag { int a; } S; int main() { struct Tag s = {5}; S s2 = s; printf(\"%d\", s2.a); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn typedef_union() {
    assert_eq!(
        run_c(
            "typedef union { int a; char c; } U; int main() { U u; u.a = 65; printf(\"%d\", u.a); return 0; }"
        ),
        vec!["65"]
    );
}
#[test]
fn typedef_enum() {
    assert_eq!(
        run_c(
            "typedef enum { A = 10, B = 20 } E; int main() { E e = B; printf(\"%d\", e); return 0; }"
        ),
        vec!["20"]
    );
}
#[test]
fn typedef_nested() {
    assert_eq!(
        run_c(
            "typedef int A; typedef A B; typedef B C; int main() { C c = 7; printf(\"%d\", c); return 0; }"
        ),
        vec!["7"]
    );
}
#[test]
fn typedef_shadowing() {
    assert_eq!(
        run_c("typedef int T; int main() { T T = 8; printf(\"%d\", T); return 0; }"),
        vec!["8"]
    );
} // T is now a variable
#[test]
fn typedef_redefinition_same() {
    assert_eq!(
        run_c("typedef int T; typedef int T; int main() { T a = 9; printf(\"%d\", a); return 0; }"),
        vec!["9"]
    );
} // C11 allows redefinition to same type
#[test]
fn typedef_block_scope() {
    assert_eq!(
        run_c("int main() { typedef int T; T a = 10; printf(\"%d\", a); return 0; }"),
        vec!["10"]
    );
}
#[test]
fn typedef_block_scope_shadows() {
    assert_eq!(
        run_c(
            "typedef int T; int main() { typedef char T; T a = 'A'; printf(\"%c\", a); return 0; }"
        ),
        vec!["A"]
    );
}
#[test]
fn typedef_const_pointer() {
    assert_eq!(
        run_c(
            "typedef int* ptr; int main() { int a = 1; const ptr p = &a; /* p is const pointer to int, not pointer to const int */ *p = 11; printf(\"%d\", *p); return 0; }"
        ),
        vec!["11"]
    );
}
#[test]
fn typedef_volatile() {
    assert_eq!(
        run_c("typedef volatile int V; int main() { V a = 12; printf(\"%d\", a); return 0; }"),
        vec!["12"]
    );
}
#[test]
fn typedef_struct_self_referential() {
    assert_eq!(
        run_c(
            "typedef struct Node { int data; struct Node* next; } Node; int main() { Node n = {13, 0}; printf(\"%d\", n.data); return 0; }"
        ),
        vec!["13"]
    );
}
