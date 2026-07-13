use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn atomic_exchange() { assert_eq!(run_c("#include <stdatomic.h>\nint main() { atomic_int val = ATOMIC_VAR_INIT(5); int old = atomic_exchange(&val, 10); printf(\"%d %d\", old, atomic_load(&val)); return 0; }"), vec!["5 10"]); }
#[test] fn atomic_compare_exchange_strong_success() { assert_eq!(run_c("#include <stdatomic.h>\nint main() { atomic_int val = ATOMIC_VAR_INIT(5); int expected = 5; _Bool res = atomic_compare_exchange_strong(&val, &expected, 10); printf(\"%d %d %d\", res, expected, atomic_load(&val)); return 0; }"), vec!["1 5 10"]); }
#[test] fn atomic_compare_exchange_strong_fail() { assert_eq!(run_c("#include <stdatomic.h>\nint main() { atomic_int val = ATOMIC_VAR_INIT(5); int expected = 3; _Bool res = atomic_compare_exchange_strong(&val, &expected, 10); printf(\"%d %d %d\", res, expected, atomic_load(&val)); return 0; }"), vec!["0 5 5"]); } // expected is updated to actual
#[test] fn atomic_compare_exchange_weak_loop() { assert_eq!(run_c("#include <stdatomic.h>\nint main() { atomic_int val = ATOMIC_VAR_INIT(5); int expected = 5; while (!atomic_compare_exchange_weak(&val, &expected, 10)) { /* weak can fail spuriously */ } printf(\"%d\", atomic_load(&val)); return 0; }"), vec!["10"]); }
