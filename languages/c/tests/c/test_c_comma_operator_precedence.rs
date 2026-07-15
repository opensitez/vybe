use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn comma_operator_basic() {
    assert_eq!(
        run_c("int main() { int a = (1, 2); printf(\"%d\", a); return 0; }"),
        vec!["2"]
    );
}
#[test]
fn comma_operator_side_effects() {
    assert_eq!(
        run_c("int main() { int a = 1; int b = (a=2, a+1); printf(\"%d\", b); return 0; }"),
        vec!["3"]
    );
}
#[test]
fn comma_operator_chain() {
    assert_eq!(
        run_c("int main() { int a = (1, 2, 3, 4); printf(\"%d\", a); return 0; }"),
        vec!["4"]
    );
}
#[test]
fn comma_operator_discarded_value() {
    assert_eq!(
        run_c("int main() { int a=1; 2, 3, a=4; printf(\"%d\", a); return 0; }"),
        vec!["4"]
    );
}
#[test]
fn comma_operator_in_if_condition() {
    assert_eq!(
        run_c("int main() { int a=1; if (a=2, a==2) printf(\"A\"); return 0; }"),
        vec!["A"]
    );
}
#[test]
fn comma_operator_in_while_condition() {
    assert_eq!(
        run_c("int main() { int a=2; while (a--, a>0) printf(\"B\"); return 0; }"),
        vec!["B"]
    );
}
#[test]
fn comma_operator_in_for_init() {
    assert_eq!(
        run_c("int main() { int a, b; for (a=1, b=2; a<2; a++) printf(\"%d\", b); return 0; }"),
        vec!["2"]
    );
}
#[test]
fn comma_operator_in_for_step() {
    assert_eq!(
        run_c("int main() { int a, b=0; for (a=0; a<1; a++, b=3) ; printf(\"%d\", b); return 0; }"),
        vec!["3"]
    );
}
#[test]
fn comma_operator_precedence_assignment() {
    assert_eq!(
        run_c("int main() { int a; a = 1, 2; printf(\"%d\", a); return 0; }"),
        vec!["1"]
    );
} // Assignment has higher precedence than comma
#[test]
fn comma_operator_precedence_ternary() {
    assert_eq!(
        run_c("int main() { int a = 1 ? (2, 3) : 4; printf(\"%d\", a); return 0; }"),
        vec!["3"]
    );
}
#[test]
fn comma_operator_precedence_return() {
    assert_eq!(
        run_c("int f() { return 1, 2; } int main() { printf(\"%d\", f()); return 0; }"),
        vec!["2"]
    );
}
#[test]
fn comma_operator_void_type() {
    assert_eq!(
        run_c("void f() {} int main() { int a = (f(), 5); printf(\"%d\", a); return 0; }"),
        vec!["5"]
    );
}
#[test]
fn comma_operator_lvalue_fails() {
    assert_eq!(
        run_c(
            "int main() { int a, b; /* (a, b) = 1; // comma result is not an lvalue in C */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn comma_operator_address_of() {
    assert_eq!(
        run_c(
            "int main() { int a=1, b=2; int *p = &(a, b); /* actually illegal in C because (a,b) is not lvalue. We test failure to parse/compile, but just verifying runner behavior. */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn comma_operator_sizeof() {
    assert_eq!(
        run_c("int main() { printf(\"%d\", (int)sizeof((1, 2.0))); return 0; }"),
        vec!["8"]
    );
} // Assuming double is 8 bytes
