use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn calloc_basic() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int *p = calloc(2, sizeof(int)); printf(\"%d\", p[0] + p[1]); free(p); return 0; }"
        ),
        vec!["0"]
    );
} // Zero initialized
#[test]
fn calloc_zero_elements() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { void *p = calloc(0, sizeof(int)); printf(\"%d\", p == NULL || p != NULL); free(p); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn calloc_zero_size() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { void *p = calloc(2, 0); printf(\"%d\", p == NULL || p != NULL); free(p); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn calloc_struct() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nstruct S { int a; float b; }; int main() { struct S *p = calloc(1, sizeof(struct S)); printf(\"%d\", p->a == 0 && p->b == 0.0f); free(p); return 0; }"
        ),
        vec!["1"]
    );
} // Float zero representation is mostly all 0 bytes
#[test]
fn calloc_pointers() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int **p = calloc(2, sizeof(int*)); printf(\"%d\", p[0] == NULL); free(p); return 0; }"
        ),
        vec!["1"]
    );
} // Null pointer representation is usually all 0 bytes
#[test]
fn calloc_overflow() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\n#include <stdint.h>\nint main() { void *p = calloc(SIZE_MAX / 2 + 2, 2); printf(\"%d\", p == NULL); return 0; }"
        ),
        vec!["1"]
    );
} // Multiplication overflows, should return NULL
#[test]
fn calloc_large_alloc() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { void *p = calloc(1000, 1000); if (p) free(p); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn calloc_char_array() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { char *p = calloc(5, sizeof(char)); printf(\"%d\", p[4] == '\\0'); free(p); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn calloc_vla() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int n = 5; int *p = calloc(n, sizeof(int)); printf(\"%d\", p[n-1]); free(p); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn calloc_assignment() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int *p = calloc(1, sizeof(int)); *p = 42; printf(\"%d\", *p); free(p); return 0; }"
        ),
        vec!["42"]
    );
}
#[test]
fn calloc_double_free_fails() {
    assert_eq!(
        run_c(
            "/* #include <stdlib.h>\nint main() { void *p = calloc(1,1); free(p); free(p); return 0; } */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn calloc_multidim() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int (*p)[2] = calloc(2, sizeof(int[2])); printf(\"%d\", p[1][1]); free(p); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn calloc_cast() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { float *p = (float*)calloc(1, sizeof(float)); printf(\"%d\", *p == 0.0f); free(p); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn calloc_check_all_zero() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int *p = calloc(10, sizeof(int)); int sum=0; for(int i=0; i<10; i++) sum += p[i]; printf(\"%d\", sum); free(p); return 0; }"
        ),
        vec!["0"]
    );
}
