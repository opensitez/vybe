use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn atomic_flag_test_and_set() { assert_eq!(run_c("#include <stdatomic.h>\nint main() { atomic_flag lock = ATOMIC_FLAG_INIT; _Bool was_set = atomic_flag_test_and_set(&lock); printf(\"%d\", was_set); return 0; }"), vec!["0"]); }
#[test] fn atomic_flag_clear() { assert_eq!(run_c("#include <stdatomic.h>\nint main() { atomic_flag lock = ATOMIC_FLAG_INIT; atomic_flag_test_and_set(&lock); atomic_flag_clear(&lock); _Bool was_set = atomic_flag_test_and_set(&lock); printf(\"%d\", was_set); return 0; }"), vec!["0"]); } // returns 0 again since it was cleared
#[test] fn atomic_flag_spinlock_simulation() { assert_eq!(run_c("#include <stdatomic.h>\nint main() { atomic_flag lock = ATOMIC_FLAG_INIT; while (atomic_flag_test_and_set_explicit(&lock, memory_order_acquire)) {} /* locked */ atomic_flag_clear_explicit(&lock, memory_order_release); printf(\"ok\"); return 0; }"), vec!["ok"]); }
