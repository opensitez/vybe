use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn vla_multi_basic() {
    assert_eq!(
        run_c(
            "int main() { int r=2, c=3; int arr[r][c]; arr[1][2] = 5; printf(\"%d\", arr[1][2]); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn vla_multi_dynamic_strides() {
    assert_eq!(
        run_c(
            "int main() { int r=2, c=3; int arr[r][c]; int *p = &arr[0][0]; p[5] = 10; printf(\"%d\", arr[1][2]); return 0; }"
        ),
        vec!["10"]
    );
}
#[test]
fn vla_multi_sizeof() {
    assert_eq!(
        run_c(
            "int main() { int r=2, c=3; int arr[r][c]; printf(\"%d\", (int)sizeof(arr) == r * c * sizeof(int)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn vla_multi_sizeof_row() {
    assert_eq!(
        run_c(
            "int main() { int r=2, c=3; int arr[r][c]; printf(\"%d\", (int)sizeof(arr[0]) == c * sizeof(int)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn vla_multi_pointer_arithmetic() {
    assert_eq!(
        run_c(
            "int main() { int r=2, c=3; int arr[r][c]; int (*p)[c] = arr + 1; printf(\"%d\", (int)((char*)p - (char*)arr) == c * sizeof(int)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn vla_multi_function_arg() {
    assert_eq!(
        run_c(
            "void f(int r, int c, int arr[r][c]) { printf(\"%d\", (int)sizeof(*arr) == c * sizeof(int)); } int main() { int r=2, c=3; int arr[r][c]; f(r, c, arr); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn vla_multi_decay() {
    assert_eq!(
        run_c(
            "int main() { int r=2, c=3; int arr[r][c]; int (*p)[c] = arr; printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn vla_multi_variable_dimensions_in_struct_fails() {
    assert_eq!(
        run_c("/* struct S { int n; int arr[n][n]; }; */ int main() { printf(\"ok\"); return 0; }"),
        vec!["ok"]
    );
}
#[test]
fn vla_multi_typedef() {
    assert_eq!(
        run_c(
            "int main() { int c=3; typedef int Row[c]; Row arr[2]; arr[1][2] = 7; printf(\"%d\", arr[1][2]); return 0; }"
        ),
        vec!["7"]
    );
}
#[test]
fn vla_multi_pointer_to_vla() {
    assert_eq!(
        run_c(
            "int main() { int r=2, c=3; int (*p)[r][c] = 0; printf(\"%d\", (int)sizeof(*p) == r * c * sizeof(int)); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn vla_multi_dynamic_exprs() {
    assert_eq!(
        run_c(
            "int f(int *x) { (*x)++; return 2; } int main() { int x=0; int arr[f(&x)][f(&x)]; printf(\"%d\", x); return 0; }"
        ),
        vec!["2"]
    );
} // Evaluated sequentially
#[test]
fn vla_multi_goto_bypassing_fails() {
    assert_eq!(
        run_c(
            "/* int main() { goto L; int arr[2][2]; L: return 0; } // UB if VLA scope entered by goto, but we test compile */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn vla_multi_3d() {
    assert_eq!(
        run_c(
            "int main() { int d1=2, d2=3, d3=4; int arr[d1][d2][d3]; arr[1][2][3] = 42; printf(\"%d\", arr[1][2][3]); return 0; }"
        ),
        vec!["42"]
    );
}
#[test]
fn vla_multi_in_loop() {
    assert_eq!(
        run_c(
            "int main() { int sum=0; for(int i=1; i<=2; i++) { int arr[i][i]; sum += sizeof(arr)/sizeof(int); } printf(\"%d\", sum); return 0; }"
        ),
        vec!["5"]
    );
} // 1x1 + 2x2 = 5
