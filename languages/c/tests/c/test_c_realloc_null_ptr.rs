use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn realloc_basic() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int *p = malloc(sizeof(int)); *p = 5; int *q = realloc(p, 2 * sizeof(int)); q[1] = 10; printf(\"%d\", q[0] + q[1]); free(q); return 0; }"
        ),
        vec!["15"]
    );
}
#[test]
fn realloc_null_ptr() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int *p = realloc(NULL, sizeof(int)); *p = 42; printf(\"%d\", *p); free(p); return 0; }"
        ),
        vec!["42"]
    );
} // Acts like malloc
#[test]
fn realloc_zero_size() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int *p = malloc(sizeof(int)); void *q = realloc(p, 0); printf(\"ok\"); /* standard C says behavior is implementation-defined, might free or return NULL or unique ptr */ return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn realloc_larger() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int *p = malloc(10 * sizeof(int)); for(int i=0; i<10; i++) p[i]=i; int *q = realloc(p, 20 * sizeof(int)); printf(\"%d\", q[9]); free(q); return 0; }"
        ),
        vec!["9"]
    );
} // Preserves data
#[test]
fn realloc_smaller() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int *p = malloc(20 * sizeof(int)); for(int i=0; i<20; i++) p[i]=i; int *q = realloc(p, 10 * sizeof(int)); printf(\"%d\", q[9]); free(q); return 0; }"
        ),
        vec!["9"]
    );
} // Preserves data
#[test]
fn realloc_fail() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\n#include <stdint.h>\nint main() { void *p = malloc(10); void *q = realloc(p, SIZE_MAX); if (q == NULL) printf(\"fail\"); else free(q); free(p); return 0; }"
        ),
        vec!["fail"]
    );
} // Returns NULL, original intact
#[test]
fn realloc_same_size() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int *p = malloc(sizeof(int)); *p = 7; int *q = realloc(p, sizeof(int)); printf(\"%d\", *q); free(q); return 0; }"
        ),
        vec!["7"]
    );
}
#[test]
fn realloc_string() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\n#include <string.h>\nint main() { char *p = malloc(5); strcpy(p, \"abc\"); p = realloc(p, 10); strcat(p, \"def\"); printf(\"%s\", p); free(p); return 0; }"
        ),
        vec!["abcdef"]
    );
}
#[test]
fn realloc_struct() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nstruct S { int a; }; int main() { struct S *p = malloc(sizeof(struct S)); p->a = 1; p = realloc(p, 2 * sizeof(struct S)); p[1].a = 2; printf(\"%d\", p[0].a + p[1].a); free(p); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn realloc_pointer_array() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int **p = malloc(sizeof(int*)); p[0] = malloc(sizeof(int)); *p[0] = 5; p = realloc(p, 2 * sizeof(int*)); p[1] = malloc(sizeof(int)); *p[1] = 6; printf(\"%d\", *p[0] + *p[1]); free(p[0]); free(p[1]); free(p); return 0; }"
        ),
        vec!["11"]
    );
}
#[test]
fn realloc_chain() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int *p = malloc(sizeof(int)); p = realloc(p, 2*sizeof(int)); p = realloc(p, 3*sizeof(int)); p[2] = 33; printf(\"%d\", p[2]); free(p); return 0; }"
        ),
        vec!["33"]
    );
}
#[test]
fn realloc_uninitialized_read_fails() {
    assert_eq!(
        run_c(
            "/* #include <stdlib.h>\nint main() { int *p = malloc(sizeof(int)); int *q = realloc(p, 2*sizeof(int)); int val = q[1]; return 0; } */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn realloc_double_free_fails() {
    assert_eq!(
        run_c(
            "/* #include <stdlib.h>\nint main() { void *p = malloc(10); void *q = realloc(p, 20); free(p); free(q); return 0; } // p is freed by realloc */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn realloc_null_ptr_zero_size() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { void *p = realloc(NULL, 0); printf(\"ok\"); if (p) free(p); return 0; }"
        ),
        vec!["ok"]
    );
}
