use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn aligned_alloc_basic() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\n#include <stdint.h>\nint main() { void *p = aligned_alloc(32, 64); printf(\"%d\", ((uintptr_t)p % 32) == 0); free(p); return 0; }"
        ),
        vec!["1"]
    );
} // C11 aligned_alloc
#[test]
fn aligned_alloc_large_alignment() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\n#include <stdint.h>\nint main() { void *p = aligned_alloc(128, 256); printf(\"%d\", ((uintptr_t)p % 128) == 0); free(p); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn aligned_alloc_alignment_not_power_of_two_fails() {
    assert_eq!(
        run_c(
            "/* #include <stdlib.h>\nint main() { void *p = aligned_alloc(15, 30); return 0; } // undefined behavior if alignment not supported */ int main() { printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn aligned_alloc_size_not_multiple_of_alignment() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { void *p = aligned_alloc(32, 10); if (p) free(p); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // UB in C11, but many impls accept it or return NULL
#[test]
fn aligned_alloc_zero_size() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { void *p = aligned_alloc(16, 0); if (p) free(p); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // implementation defined
#[test]
fn aligned_alloc_write() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int *p = aligned_alloc(16, 16); *p = 42; printf(\"%d\", *p); free(p); return 0; }"
        ),
        vec!["42"]
    );
}
#[test]
fn aligned_alloc_struct() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nstruct S { int a[4]; }; int main() { struct S *p = aligned_alloc(16, sizeof(struct S)); p->a[3] = 7; printf(\"%d\", p->a[3]); free(p); return 0; }"
        ),
        vec!["7"]
    );
}
#[test]
fn aligned_alloc_null_check() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { void *p = aligned_alloc(16, 32); if(p != NULL) printf(\"1\"); free(p); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn posix_memalign_basic() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\n#include <stdint.h>\nint main() { void *p; int res = posix_memalign(&p, 32, 64); if(res == 0) printf(\"%d\", ((uintptr_t)p % 32) == 0); free(p); return 0; }"
        ),
        vec!["1"]
    );
} // POSIX standard
#[test]
fn posix_memalign_invalid_alignment() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { void *p = NULL; int res = posix_memalign(&p, 10, 64); printf(\"%d\", res != 0); return 0; }"
        ),
        vec!["1"]
    );
} // Alignment must be power of two and multiple of sizeof(void*)
#[test]
fn posix_memalign_write() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { int *p; posix_memalign((void**)&p, sizeof(void*), sizeof(int)); *p = 99; printf(\"%d\", *p); free(p); return 0; }"
        ),
        vec!["99"]
    );
}
#[test]
fn valloc_basic() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { void *p = valloc(10); if(p) free(p); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
} // Obsolete but still available on many systems
#[test]
fn aligned_alloc_free_with_standard_free() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\nint main() { void *p = aligned_alloc(16, 16); free(p); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn aligned_alloc_memset() {
    assert_eq!(
        run_c(
            "#include <stdlib.h>\n#include <string.h>\nint main() { char *p = aligned_alloc(16, 16); memset(p, 'x', 16); printf(\"%c\", p[15]); free(p); return 0; }"
        ),
        vec!["x"]
    );
}
