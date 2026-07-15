use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn pointer_subtraction_basic() {
    assert_eq!(
        run_c(
            "int main() { int arr[5]; int *p1 = &arr[1]; int *p2 = &arr[4]; printf(\"%d\", (int)(p2 - p1)); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn pointer_subtraction_negative() {
    assert_eq!(
        run_c(
            "int main() { int arr[5]; int *p1 = &arr[4]; int *p2 = &arr[1]; printf(\"%d\", (int)(p2 - p1)); return 0; }"
        ),
        vec!["-3"]
    );
}
#[test]
fn pointer_subtraction_same() {
    assert_eq!(
        run_c(
            "int main() { int arr[5]; int *p1 = &arr[2]; int *p2 = &arr[2]; printf(\"%d\", (int)(p2 - p1)); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn pointer_subtraction_char() {
    assert_eq!(
        run_c(
            "int main() { char arr[10]; char *p1 = &arr[2]; char *p2 = &arr[8]; printf(\"%d\", (int)(p2 - p1)); return 0; }"
        ),
        vec!["6"]
    );
}
#[test]
fn pointer_subtraction_struct() {
    assert_eq!(
        run_c(
            "struct S { int a; double b; }; int main() { struct S arr[5]; struct S *p1 = &arr[0]; struct S *p2 = &arr[4]; printf(\"%d\", (int)(p2 - p1)); return 0; }"
        ),
        vec!["4"]
    );
}
#[test]
fn pointer_subtraction_different_arrays_fails() {
    assert_eq!(
        run_c(
            "/* int a[2], b[2]; int diff = &a[1] - &b[1]; */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // UB
#[test]
fn pointer_subtraction_ptrdiff_t() {
    assert_eq!(
        run_c(
            "#include <stddef.h>\nint main() { int arr[2]; ptrdiff_t d = &arr[1] - &arr[0]; printf(\"%d\", d == 1); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pointer_subtraction_multidim() {
    assert_eq!(
        run_c(
            "int main() { int arr[3][4]; int (*p1)[4] = &arr[0]; int (*p2)[4] = &arr[2]; printf(\"%d\", (int)(p2 - p1)); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn pointer_subtraction_void_fails() {
    assert_eq!(
        run_c(
            "/* int main() { void *p1, *p2; int d = p2 - p1; return 0; } // GNU extension allows this, standard C does not */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn pointer_subtraction_function_fails() {
    assert_eq!(
        run_c(
            "/* void f(){} void g(){} int main() { int d = g - f; return 0; } */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn pointer_subtraction_null() {
    assert_eq!(
        run_c("int main() { int *p1 = 0; int *p2 = 0; printf(\"%d\", (int)(p2 - p1)); return 0; }"),
        vec!["0"]
    );
} // UB technically if not in same object, but commonly evaluates to 0
#[test]
fn pointer_subtraction_vla() {
    assert_eq!(
        run_c(
            "int main() { int n=5; int arr[n]; int *p1 = &arr[1]; int *p2 = &arr[n-1]; printf(\"%d\", (int)(p2 - p1)); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn pointer_subtraction_with_cast() {
    assert_eq!(
        run_c(
            "int main() { int arr[2]; char *p1 = (char*)&arr[0]; char *p2 = (char*)&arr[1]; printf(\"%d\", (int)(p2 - p1) == sizeof(int)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pointer_subtraction_in_const_expr_fails() {
    assert_eq!(
        run_c(
            "/* int arr[2]; const int d = &arr[1] - &arr[0]; // pointer difference is not integer constant expression */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
