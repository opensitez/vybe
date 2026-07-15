use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn atomic_int_init_load() {
    assert_eq!(
        run_c(
            "#include <stdatomic.h>\nint main() { atomic_int val = ATOMIC_VAR_INIT(42); printf(\"%d\", atomic_load(&val)); return 0; }"
        ),
        vec!["42"]
    );
}
#[test]
fn atomic_store_load() {
    assert_eq!(
        run_c(
            "#include <stdatomic.h>\nint main() { atomic_long val; atomic_store(&val, 123L); printf(\"%ld\", atomic_load(&val)); return 0; }"
        ),
        vec!["123"]
    );
}
#[test]
fn atomic_explicit_memory_order() {
    assert_eq!(
        run_c(
            "#include <stdatomic.h>\nint main() { atomic_int val = 0; atomic_store_explicit(&val, 5, memory_order_release); printf(\"%d\", atomic_load_explicit(&val, memory_order_acquire)); return 0; }"
        ),
        vec!["5"]
    );
}
#[test]
fn atomic_pointer_type() {
    assert_eq!(
        run_c(
            "#include <stdatomic.h>\nint main() { int target = 10; _Atomic(int*) ptr; atomic_init(&ptr, &target); int *p = atomic_load(&ptr); printf(\"%d\", *p); return 0; }"
        ),
        vec!["10"]
    );
}
#[test]
fn atomic_is_lock_free() {
    assert_eq!(
        run_c(
            "#include <stdatomic.h>\nint main() { atomic_int val; printf(\"%d\", atomic_is_lock_free(&val) >= 0); return 0; }"
        ),
        vec!["1"]
    );
} // It returns boolean-like
