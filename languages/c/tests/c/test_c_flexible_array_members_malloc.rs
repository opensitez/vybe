use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn fam_malloc_basic() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nstruct S { int len; int data[]; }; int main() { struct S *p = malloc(sizeof(struct S) + 5 * sizeof(int)); p->len = 5; p->data[4] = 42; printf(\"%d\", p->data[4]); free(p); return 0; }"
        ),
        vec!["42"]
    );
}
#[test]
fn fam_malloc_zero_length() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nstruct S { int len; int data[]; }; int main() { struct S *p = malloc(sizeof(struct S)); p->len = 0; printf(\"%d\", p->len); free(p); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn fam_malloc_loop() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nstruct S { int len; int data[]; }; int main() { struct S *p = malloc(sizeof(struct S) + 3 * sizeof(int)); int sum=0; for(int i=0; i<3; i++) { p->data[i] = i; sum += p->data[i]; } printf(\"%d\", sum); free(p); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn fam_realloc() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nstruct S { int len; int data[]; }; int main() { struct S *p = malloc(sizeof(struct S) + sizeof(int)); p->data[0] = 1; p = realloc(p, sizeof(struct S) + 2 * sizeof(int)); p->data[1] = 2; printf(\"%d\", p->data[0] + p->data[1]); free(p); return 0; }"
        ),
        vec!["3"]
    );
}
#[test]
fn fam_calloc() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nstruct S { int len; int data[]; }; int main() { struct S *p = calloc(1, sizeof(struct S) + 2 * sizeof(int)); printf(\"%d\", p->data[1]); free(p); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn fam_pointer_arithmetic() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nstruct S { int len; int data[]; }; int main() { struct S *p = malloc(sizeof(struct S) + 2 * sizeof(int)); int *d = p->data; d[1] = 9; printf(\"%d\", p->data[1]); free(p); return 0; }"
        ),
        vec!["9"]
    );
}
#[test]
fn fam_nested_pointers() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nstruct S { int len; int *data[]; }; int main() { struct S *p = malloc(sizeof(struct S) + 2 * sizeof(int*)); int x=5; p->data[1] = &x; printf(\"%d\", *p->data[1]); free(p); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn fam_offsetof() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\n#include <stddef.h>\nstruct S { char c; int data[]; }; int main() { struct S *p = malloc(sizeof(struct S) + sizeof(int)); p->data[0] = 7; int *d = (int*)((char*)p + offsetof(struct S, data)); printf(\"%d\", d[0]); free(p); return 0; }"
        ),
        vec!["7"]
    );
}
#[test]
// A pointer to a struct with a flexible array member is an ORDINARY object
// pointer — that is the property worth asserting. Comparing `sizeof(p)` against
// a literal 4 asserted the target's pointer width instead: true on wasm32,
// false under `cc` on any 64-bit host, so the test could never agree with the
// reference compiler.
fn fam_sizeof_pointer() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nstruct S { int len; int data[]; }; int main() { struct S *p; printf(\"%d\", (int)(sizeof(p) == sizeof(void*))); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn fam_sizeof_deref() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nstruct S { int len; int data[]; }; int main() { struct S *p; printf(\"%d\", (int)sizeof(*p) == sizeof(struct S)); return 0; }"
        ),
        vec!["1"]
    );
} // sizeof(*p) is just the declared members
#[test]
fn fam_memcpy() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\n#include <string.h>\nstruct S { int len; int data[]; }; int main() { struct S *p1 = malloc(sizeof(struct S) + 2 * sizeof(int)); p1->len=2; p1->data[0]=1; p1->data[1]=2; struct S *p2 = malloc(sizeof(struct S) + 2 * sizeof(int)); memcpy(p2, p1, sizeof(struct S) + 2 * sizeof(int)); printf(\"%d\", p2->data[1]); free(p1); free(p2); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn fam_free() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nstruct S { int len; int data[]; }; int main() { struct S *p = malloc(sizeof(struct S) + 10); free(p); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn fam_pass_pointer() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nstruct S { int len; int data[]; }; void f(struct S *p) { printf(\"%d\", p->data[0]); } int main() { struct S *p = malloc(sizeof(struct S) + sizeof(int)); p->data[0] = 88; f(p); free(p); return 0; }"
        ),
        vec!["88"]
    );
}
#[test]
fn fam_dynamic_allocation_in_function() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nstruct S { int len; int data[]; }; struct S *create(int n) { struct S *p = malloc(sizeof(struct S) + n * sizeof(int)); p->len = n; return p; } int main() { struct S *p = create(2); p->data[1] = 4; printf(\"%d\", p->data[1]); free(p); return 0; }"
        ),
        vec!["4"]
    );
}
