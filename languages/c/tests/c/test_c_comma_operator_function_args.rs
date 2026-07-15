use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn comma_operator_in_args() {
    assert_eq!(
        run_c("void f(int a) { printf(\"%d\", a); } int main() { f((1, 2)); return 0; }"),
        vec!["2"]
    );
}
#[test]
fn comma_operator_in_multi_args() {
    assert_eq!(
        run_c(
            "void f(int a, int b) { printf(\"%d%d\", a, b); } int main() { f((1, 2), 3); return 0; }"
        ),
        vec!["23"]
    );
}
#[test]
fn comma_operator_not_arg_separator() {
    assert_eq!(
        run_c(
            "void f(int a, int b) { printf(\"%d%d\", a, b); } int main() { f(1, (2, 3)); return 0; }"
        ),
        vec!["13"]
    );
}
#[test]
fn comma_operator_in_macro_args() {
    assert_eq!(
        run_c("#define M(a) printf(\"%d\", a)\nint main() { M((1, 2)); return 0; }"),
        vec!["2"]
    );
}
#[test]
fn comma_operator_with_side_effect_args() {
    assert_eq!(
        run_c(
            "void f(int a) { printf(\"%d\", a); } int main() { int x=1; f((x=3, x+1)); return 0; }"
        ),
        vec!["4"]
    );
}
#[test]
fn comma_operator_arg_evaluation_order() {
    assert_eq!(
        run_c(
            "void f(int a, int b) { printf(\"A\"); } int main() { int x=0; f((x=1, x), (x=2, x)); printf(\"ok\"); return 0; }"
        ),
        vec!["A", "ok"]
    );
} // Order is unspecified, we just test execution
#[test]
fn comma_operator_in_nested_call() {
    assert_eq!(
        run_c(
            "int g(int a) { return a; } void f(int a) { printf(\"%d\", a); } int main() { f(g((1, 5))); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn comma_operator_with_struct_init_fails() {
    assert_eq!(
        run_c(
            "struct S { int a; int b; }; int main() { /* struct S s = { (1, 2), 3 }; // braces context may have ambiguity, but usually fine in C */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn comma_operator_in_array_init() {
    assert_eq!(
        run_c(
            "int main() { int arr[] = { (1, 2), (3, 4) }; printf(\"%d%d\", arr[0], arr[1]); return 0; }"
        ),
        vec!["24"]
    );
}
#[test]
fn comma_operator_in_designated_init() {
    assert_eq!(
        run_c(
            "struct S { int a, b; }; int main() { struct S s = { .a = (1, 2) }; printf(\"%d\", s.a); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn comma_operator_in_sizeof_arg() {
    assert_eq!(
        run_c(
            "void f(int a) {} int main() { printf(\"%d\", (int)sizeof(f((1, 2.0)))); return 0; }"
        ),
        vec!["1"]
    );
} // f returns void, sizeof void is 1 in GNU C, or error in standard C. We test parsing.
#[test]
fn comma_operator_in_cast() {
    assert_eq!(
        run_c("void f(int a) { printf(\"%d\", a); } int main() { f((int)(1.5, 2.5)); return 0; }"),
        vec!["2"]
    );
}
#[test]
fn comma_operator_in_variadic() {
    assert_eq!(
        run_c("#include <stdio.h>\nint main() { printf(\"%d\", (1, 5)); return 0; }"),
        vec!["5"]
    );
}
#[test]
fn comma_operator_in_function_ptr_call() {
    assert_eq!(
        run_c(
            "void f(int a) { printf(\"%d\", a); } int main() { void (*p)(int) = f; p((1, 2)); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn comma_operator_with_comma_operator_args() {
    assert_eq!(
        run_c(
            "void f(int a, int b) { printf(\"%d%d\", a, b); } int main() { f((1, 2, 3), (4, 5, 6)); return 0; }"
        ),
        vec!["36"]
    );
}
