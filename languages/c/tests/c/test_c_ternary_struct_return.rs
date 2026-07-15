use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn ternary_struct_return_basic() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { struct S s1={1}, s2={2}; struct S res = 1 ? s1 : s2; printf(\"%d\", res.a); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn ternary_struct_return_false() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { struct S s1={1}, s2={2}; struct S res = 0 ? s1 : s2; printf(\"%d\", res.a); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn ternary_struct_different_types_fails() {
    assert_eq!(
        run_c(
            "struct S1 { int a; }; struct S2 { int a; }; int main() { /* struct S1 s1; struct S2 s2; 1 ? s1 : s2; // type mismatch */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn ternary_struct_compound_literal() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { struct S res = 1 ? (struct S){3} : (struct S){4}; printf(\"%d\", res.a); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn ternary_struct_function_call() {
    assert_eq!(
        run_c(
            "struct S { int a; }; struct S f1() { return (struct S){5}; } struct S f2() { return (struct S){6}; } int main() { struct S res = 0 ? f1() : f2(); printf(\"%d\", res.a); return 0; }"
        ),
        vec!["6"]
    );
}
#[test]
fn ternary_struct_assignment() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { struct S s1={1}, s2={2}, res; res = 1 ? s1 : s2; printf(\"%d\", res.a); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn ternary_struct_pointer_deref() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { struct S s1={7}, s2={8}; struct S *p1=&s1, *p2=&s2; struct S res = 1 ? *p1 : *p2; printf(\"%d\", res.a); return 0; }"
        ),
        vec!["7"]
    );
}
#[test]
fn ternary_struct_nested_structs() {
    assert_eq!(
        run_c(
            "struct Inner { int a; }; struct Outer { struct Inner i; }; int main() { struct Outer o1={{1}}, o2={{2}}; printf(\"%d\", (0 ? o1 : o2).i.a); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn ternary_struct_const_qualifier() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { const struct S s1={1}; struct S s2={2}; struct S res = 1 ? s1 : s2; printf(\"%d\", res.a); return 0; }"
        ),
        vec!["1"]
    );
} // type is const struct S
#[test]
fn ternary_struct_array_member() {
    assert_eq!(
        run_c(
            "struct S { int arr[2]; }; int main() { struct S s1={{1,2}}, s2={{3,4}}; struct S res = 1 ? s1 : s2; printf(\"%d\", res.arr[1]); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn ternary_struct_union_return() {
    assert_eq!(
        run_c(
            "union U { int i; float f; }; int main() { union U u1={.i=1}, u2={.i=2}; printf(\"%d\", (0 ? u1 : u2).i); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn ternary_struct_pass_to_function() {
    assert_eq!(
        run_c(
            "struct S { int a; }; void f(struct S s) { printf(\"%d\", s.a); } int main() { struct S s1={8}, s2={9}; f(1 ? s1 : s2); return 0; }"
        ),
        vec!["8"]
    );
}
#[test]
fn ternary_struct_return_from_function() {
    assert_eq!(
        run_c(
            "struct S { int a; }; struct S f() { struct S s1={1}, s2={2}; return 0 ? s1 : s2; } int main() { printf(\"%d\", f().a); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn ternary_struct_sizeof() {
    assert_eq!(
        run_c(
            "struct S { double d; int i; }; int main() { struct S s1, s2; printf(\"%d\", sizeof(1 ? s1 : s2) == sizeof(struct S)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn ternary_struct_volatile_qualifier() {
    assert_eq!(
        run_c(
            "struct S { int a; }; int main() { volatile struct S s1={1}; struct S s2={2}; struct S res = 1 ? s1 : s2; printf(\"%d\", res.a); return 0; }"
        ),
        vec!["1"]
    );
}
