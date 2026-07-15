use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn atomic_fetch_add() {
    assert_eq!(
        run_c(
            "#include <stdatomic.h>\nint main() { atomic_int val = ATOMIC_VAR_INIT(10); int old = atomic_fetch_add(&val, 5); printf(\"%d %d\", old, atomic_load(&val)); return 0; }"
        ),
        vec!["10 15"]
    );
}
#[test]
fn atomic_fetch_sub() {
    assert_eq!(
        run_c(
            "#include <stdatomic.h>\nint main() { atomic_int val = ATOMIC_VAR_INIT(10); int old = atomic_fetch_sub(&val, 3); printf(\"%d %d\", old, atomic_load(&val)); return 0; }"
        ),
        vec!["10 7"]
    );
}
#[test]
fn atomic_fetch_or() {
    assert_eq!(
        run_c(
            "#include <stdatomic.h>\nint main() { atomic_int val = ATOMIC_VAR_INIT(0x10); atomic_fetch_or(&val, 0x01); printf(\"%x\", atomic_load(&val)); return 0; }"
        ),
        vec!["11"]
    );
}
#[test]
fn atomic_fetch_and() {
    assert_eq!(
        run_c(
            "#include <stdatomic.h>\nint main() { atomic_int val = ATOMIC_VAR_INIT(0x13); atomic_fetch_and(&val, 0x10); printf(\"%x\", atomic_load(&val)); return 0; }"
        ),
        vec!["10"]
    );
}
#[test]
fn atomic_fetch_xor() {
    assert_eq!(
        run_c(
            "#include <stdatomic.h>\nint main() { atomic_int val = ATOMIC_VAR_INIT(0x11); atomic_fetch_xor(&val, 0x01); printf(\"%x\", atomic_load(&val)); return 0; }"
        ),
        vec!["10"]
    );
}
#[test]
fn atomic_operator_overloads() {
    assert_eq!(
        run_c(
            "#include <stdatomic.h>\nint main() { _Atomic int val = 5; val += 2; val++; printf(\"%d\", val); return 0; }"
        ),
        vec!["8"]
    );
}
