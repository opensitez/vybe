use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn mtx_plain_lock_unlock() {
    assert_eq!(
        run_c(
            "#include <threads.h>\nint main() { mtx_t mtx; if (mtx_init(&mtx, mtx_plain) == thrd_success) { mtx_lock(&mtx); mtx_unlock(&mtx); mtx_destroy(&mtx); printf(\"ok\"); } return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn mtx_recursive_lock_twice() {
    assert_eq!(
        run_c(
            "#include <threads.h>\nint main() { mtx_t mtx; if (mtx_init(&mtx, mtx_plain | mtx_recursive) == thrd_success) { mtx_lock(&mtx); mtx_lock(&mtx); mtx_unlock(&mtx); mtx_unlock(&mtx); mtx_destroy(&mtx); printf(\"ok\"); } return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn mtx_trylock_success() {
    assert_eq!(
        run_c(
            "#include <threads.h>\nint main() { mtx_t mtx; if (mtx_init(&mtx, mtx_plain) == thrd_success) { if (mtx_trylock(&mtx) == thrd_success) { mtx_unlock(&mtx); printf(\"ok\"); } mtx_destroy(&mtx); } return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn mtx_timed_init_compile() {
    assert_eq!(
        run_c(
            "#include <threads.h>\nint main() { mtx_t mtx; if (mtx_init(&mtx, mtx_timed) == thrd_success) { mtx_destroy(&mtx); printf(\"ok\"); } return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn mtx_timedlock_success() {
    assert_eq!(
        run_c(
            "#include <threads.h>\n#include <time.h>\nint main() { mtx_t mtx; if (mtx_init(&mtx, mtx_timed) == thrd_success) { struct timespec ts; timespec_get(&ts, TIME_UTC); ts.tv_sec += 1; if (mtx_timedlock(&mtx, &ts) == thrd_success) { mtx_unlock(&mtx); printf(\"ok\"); } mtx_destroy(&mtx); } return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn threads_multiple_workers_mtx() {
    assert_eq!(
        run_c(
            "#include <threads.h>\nmtx_t m; int counter = 0;\nint worker(void *arg) { mtx_lock(&m); counter++; mtx_unlock(&m); return 0; }\nint main() { mtx_init(&m, mtx_plain); thrd_t t1, t2; thrd_create(&t1, worker, NULL); thrd_create(&t2, worker, NULL); thrd_join(t1, NULL); thrd_join(t2, NULL); printf(\"%d\", counter); mtx_destroy(&m); return 0; }"
        ),
        vec!["2"]
    );
}
