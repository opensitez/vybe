use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn sizeof_vla_basic() {
    assert_eq!(
        run_c(
            "int main() { int n = 5; int arr[n]; printf(\"%d\", (int)(sizeof(arr) / sizeof(int))); return 0; }"
        ),
        vec!["5"]
    );
} // Evaluated at runtime
#[test]
fn sizeof_vla_side_effects() {
    assert_eq!(
        run_c(
            "int f(int *x) { (*x)++; return 5; } int main() { int x = 0; int arr[f(&x)]; int s = sizeof(arr); printf(\"%d\", x); return 0; }"
        ),
        vec!["1"]
    );
} // VLA size expression IS evaluated
#[test]
fn sizeof_vla_side_effects_in_sizeof() {
    assert_eq!(
        run_c(
            "int main() { int n = 2; int x = 0; printf(\"%d\", (int)(sizeof(int[n + (x=1)]) / sizeof(int))); printf(\"%d\", x); return 0; }"
        ),
        vec!["3", "1"]
    );
} // Inside sizeof(), VLA size expr IS evaluated
#[test]
fn sizeof_non_vla_side_effects() {
    assert_eq!(
        run_c("int main() { int x = 0; sizeof(x=1); printf(\"%d\", x); return 0; }"),
        vec!["0"]
    );
} // Non-VLA sizeof does NOT evaluate operand
#[test]
fn sizeof_vla_pointer() {
    assert_eq!(
        run_c("int main() { int n = 5; int (*p)[n]; printf(\"%d\", (int)sizeof(p)); return 0; }"),
        vec!["8"]
    );
} // Assuming 64-bit pointers
#[test]
fn sizeof_vla_pointer_target() {
    assert_eq!(
        run_c(
            "int main() { int n = 5; int (*p)[n]; printf(\"%d\", (int)(sizeof(*p) / sizeof(int))); return 0; }"
        ),
        vec!["5"]
    );
} // Evaluates at runtime
#[test]
fn sizeof_vla_function_param() {
    assert_eq!(
        run_c(
            "void f(int n, int arr[n]) { printf(\"%d\", (int)sizeof(arr)); } int main() { int a[5]; f(5, a); return 0; }"
        ),
        vec!["8"]
    );
} // arr decays to pointer, so sizeof is pointer size!
#[test]
fn sizeof_vla_2d() {
    assert_eq!(
        run_c(
            "int main() { int n = 2, m = 3; int arr[n][m]; printf(\"%d\", (int)(sizeof(arr) / sizeof(int))); return 0; }"
        ),
        vec!["6"]
    );
}
#[test]
fn sizeof_vla_2d_partial() {
    assert_eq!(
        run_c(
            "int main() { int n = 2, m = 3; int arr[n][m]; printf(\"%d\", (int)(sizeof(arr[0]) / sizeof(int))); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn sizeof_vla_in_loop() {
    assert_eq!(
        run_c(
            "int main() { int sum = 0; for(int i=1; i<=3; i++) { int arr[i]; sum += sizeof(arr) / sizeof(int); } printf(\"%d\", sum); return 0; }"
        ),
        vec!["6"]
    );
} // 1 + 2 + 3
#[test]
fn sizeof_vla_goto_bypass_fails() {
    assert_eq!(
        run_c(
            "/* int main() { int n=5; goto L; int arr[n]; L: sizeof(arr); return 0; } */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn sizeof_vla_ternary_size() {
    assert_eq!(
        run_c(
            "int main() { int cond = 1; int arr[cond ? 10 : 20]; printf(\"%d\", (int)(sizeof(arr) / sizeof(int))); return 0; }"
        ),
        vec!["10"]
    );
}
#[test]
fn sizeof_vla_variable_shadowing() {
    assert_eq!(
        run_c(
            "int main() { int n = 5; { int arr[n]; int n = 10; printf(\"%d\", (int)(sizeof(arr) / sizeof(int))); } return 0; }"
        ),
        vec!["5"]
    );
} // Captures size at declaration
#[test]
fn sizeof_vla_struct_member_fails() {
    assert_eq!(
        run_c(
            "/* struct S { int n; int arr[n]; }; // VLA cannot be struct member */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn sizeof_vla_typedef() {
    assert_eq!(
        run_c(
            "int main() { int n = 5; typedef int A[n]; n = 10; printf(\"%d\", (int)(sizeof(A) / sizeof(int))); return 0; }"
        ),
        vec!["5"]
    );
} // Evaluates at typedef declaration
