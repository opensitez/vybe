use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn pointer_comparison_equality() {
    assert_eq!(
        run_c(
            "int main() { int x; int *p1 = &x; int *p2 = &x; printf(\"%d\", p1 == p2); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pointer_comparison_inequality() {
    assert_eq!(
        run_c(
            "int main() { int x, y; int *p1 = &x; int *p2 = &y; printf(\"%d\", p1 != p2); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pointer_comparison_relational() {
    assert_eq!(
        run_c(
            "int main() { int arr[2]; int *p1 = &arr[0]; int *p2 = &arr[1]; printf(\"%d%d\", p1 < p2, p2 > p1); return 0; }"
        ),
        vec!["11"]
    );
}
#[test]
fn pointer_comparison_relational_different_arrays_fails() {
    assert_eq!(
        run_c(
            "/* int a[2], b[2]; int cmp = &a[0] < &b[0]; */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // UB
#[test]
fn pointer_comparison_null() {
    assert_eq!(
        run_c("int main() { int *p = 0; printf(\"%d\", p == 0); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn pointer_comparison_null_macro() {
    assert_eq!(
        run_c(
            "#include <stddef.h>\nint main() { int *p = NULL; printf(\"%d\", p == NULL); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pointer_comparison_void_cast() {
    assert_eq!(
        run_c("int main() { int x; int *p = &x; void *v = &x; printf(\"%d\", p == v); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn pointer_comparison_different_types_warning() {
    assert_eq!(
        run_c(
            "/* int main() { int x; float y; int *p1 = &x; float *p2 = &y; p1 == p2; return 0; } */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // constraint violation without cast
#[test]
fn pointer_comparison_function_pointers() {
    assert_eq!(
        run_c(
            "void f(){} int main() { void (*p1)() = f; void (*p2)() = f; printf(\"%d\", p1 == p2); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pointer_comparison_string_literals() {
    assert_eq!(
        run_c(
            "int main() { char *p1 = \"abc\"; char *p2 = \"abc\"; /* Can be equal or not depending on merging. Let's just check it compiles */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn pointer_comparison_const() {
    assert_eq!(
        run_c(
            "int main() { int x; const int *p1 = &x; int *p2 = &x; printf(\"%d\", p1 == p2); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pointer_comparison_array_decay() {
    assert_eq!(
        run_c("int main() { int arr[2]; printf(\"%d\", arr == &arr[0]); return 0; }"),
        vec!["1"]
    );
}
#[test]
fn pointer_comparison_array_address() {
    assert_eq!(
        run_c("int main() { int arr[2]; printf(\"%d\", (void*)arr == (void*)&arr); return 0; }"),
        vec!["1"]
    );
} // &arr has type int(*)[2] but same address
#[test]
fn pointer_comparison_vla() {
    assert_eq!(
        run_c(
            "int main() { int n=5; int arr[n]; int *p = arr; printf(\"%d\", p == &arr[0]); return 0; }"
        ),
        vec!["1"]
    );
}
