use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn thrd_create_join() {
    assert_eq!(
        run_c(
            "#include <threads.h>\nint worker(void *arg) { return 42; }\nint main() { thrd_t t; if (thrd_create(&t, worker, NULL) == thrd_success) { int res; thrd_join(t, &res); printf(\"%d\", res); } return 0; }"
        ),
        vec!["42"]
    );
}
#[test]
fn thrd_create_detach() {
    assert_eq!(
        run_c(
            "#include <threads.h>\nint worker(void *arg) { return 0; }\nint main() { thrd_t t; if (thrd_create(&t, worker, NULL) == thrd_success) { if (thrd_detach(t) == thrd_success) printf(\"ok\"); } return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn thrd_create_arg() {
    assert_eq!(
        run_c(
            "#include <threads.h>\nint worker(void *arg) { int *val = arg; return *val * 2; }\nint main() { thrd_t t; int val = 5; if (thrd_create(&t, worker, &val) == thrd_success) { int res; thrd_join(t, &res); printf(\"%d\", res); } return 0; }"
        ),
        vec!["10"]
    );
}
#[test]
fn thrd_current_equal() {
    assert_eq!(
        run_c(
            "#include <threads.h>\nint main() { thrd_t t = thrd_current(); printf(\"%d\", thrd_equal(t, thrd_current()) != 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn thrd_sleep_basic() {
    assert_eq!(
        run_c(
            "#include <threads.h>\n#include <time.h>\nint main() { struct timespec duration; duration.tv_sec = 0; duration.tv_nsec = 1000000; /* 1ms */ thrd_sleep(&duration, NULL); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn thrd_yield_basic() {
    assert_eq!(
        run_c("#include <threads.h>\nint main() { thrd_yield(); printf(\"ok\"); return 0; }"),
        vec!["ok"]
    );
}
#[test]
fn thrd_exit_no_return() {
    assert_eq!(
        run_c(
            "#include <threads.h>\nint worker(void *arg) { thrd_exit(99); return 0; }\nint main() { thrd_t t; thrd_create(&t, worker, NULL); int res; thrd_join(t, &res); printf(\"%d\", res); return 0; }"
        ),
        vec!["99"]
    );
}
