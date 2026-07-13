use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn call_once_single_thread() { assert_eq!(run_c("#include <threads.h>\nonce_flag flag = ONCE_FLAG_INIT;\nint counter = 0;\nvoid init(void) { counter++; }\nint main() { call_once(&flag, init); call_once(&flag, init); printf(\"%d\", counter); return 0; }"), vec!["1"]); }
#[test] fn call_once_multi_thread() { assert_eq!(run_c("#include <threads.h>\nonce_flag flag = ONCE_FLAG_INIT;\nint counter = 0;\nvoid init(void) { counter++; }\nint worker(void *arg) { call_once(&flag, init); return 0; }\nint main() { thrd_t t1, t2; thrd_create(&t1, worker, NULL); thrd_create(&t2, worker, NULL); thrd_join(t1, NULL); thrd_join(t2, NULL); printf(\"%d\", counter); return 0; }"), vec!["1"]); }
