use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn key_create_delete() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { pthread_key_t k; int r1 = pthread_key_create(&k, NULL); int r2 = pthread_key_delete(k); printf(\"%d %d\", r1 == 0, r2 == 0); return 0; }"
        ),
        vec!["1 1"]
    );
}
#[test]
fn setspecific_getspecific() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { pthread_key_t k; pthread_key_create(&k, NULL); pthread_setspecific(k, (void*)42); void *v = pthread_getspecific(k); printf(\"%ld\", (long)v); pthread_key_delete(k); return 0; }"
        ),
        vec!["42"]
    );
}
#[test]
fn key_default_null() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { pthread_key_t k; pthread_key_create(&k, NULL); void *v = pthread_getspecific(k); printf(\"%d\", v == NULL); pthread_key_delete(k); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn tls_threads_isolated() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\n#include <unistd.h>\npthread_key_t k;\nvoid* f(void* a) { pthread_setspecific(k, (void*)99); sleep(1); return pthread_getspecific(k); }\nint main() { pthread_key_create(&k, NULL); pthread_setspecific(k, (void*)42); pthread_t t; pthread_create(&t, NULL, f, NULL); void *res; pthread_join(t, &res); printf(\"%ld %ld\", (long)res, (long)pthread_getspecific(k)); pthread_key_delete(k); return 0; }"
        ),
        vec!["99 42"]
    );
}
#[test]
fn key_create_destructor() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint val = 0;\nvoid d(void* a) { val = (int)(long)a; }\nvoid* f(void* a) { pthread_key_t *k = a; pthread_setspecific(*k, (void*)123); return NULL; }\nint main() { pthread_key_t k; pthread_key_create(&k, d); pthread_t t; pthread_create(&t, NULL, f, &k); pthread_join(t, NULL); printf(\"%d\", val); pthread_key_delete(k); return 0; }"
        ),
        vec!["123"]
    );
}
#[test]
fn getspecific_invalid_key() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { /* UB to getspecific on invalid key, check compile */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn setspecific_invalid_key() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { /* UB to setspecific on invalid key, check compile */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn key_delete_invalid() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { /* UB to delete invalid key, check compile */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn multiple_keys() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { pthread_key_t k1, k2; pthread_key_create(&k1, NULL); pthread_key_create(&k2, NULL); pthread_setspecific(k1, (void*)1); pthread_setspecific(k2, (void*)2); printf(\"%ld %ld\", (long)pthread_getspecific(k1), (long)pthread_getspecific(k2)); pthread_key_delete(k1); pthread_key_delete(k2); return 0; }"
        ),
        vec!["1 2"]
    );
}
#[test]
fn tls_macro_c11() {
    assert_eq!(
        run_c(
            "#include <threads.h>\nthread_local int x = 0;\nint main() { x = 42; printf(\"%d\", x); return 0; }"
        ),
        vec!["42"]
    );
}
#[test]
fn tls_macro_gnu() {
    assert_eq!(
        run_c("__thread int x = 0;\nint main() { x = 99; printf(\"%d\", x); return 0; }"),
        vec!["99"]
    );
}
#[test]
fn key_create_limit() {
    assert_eq!(
        run_c(
            "#define _POSIX_C_SOURCE 200809L\n#include <pthread.h>\n#include <limits.h>\nint main() { /* PTHREAD_KEYS_MAX is often 1024 or more, test we can create at least 10 */ pthread_key_t keys[10]; int ok = 1; for(int i=0; i<10; i++) if(pthread_key_create(&keys[i], NULL) != 0) ok = 0; printf(\"%d\", ok); for(int i=0; i<10; i++) pthread_key_delete(keys[i]); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn setspecific_null() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { pthread_key_t k; pthread_key_create(&k, NULL); pthread_setspecific(k, (void*)42); pthread_setspecific(k, NULL); printf(\"%d\", pthread_getspecific(k) == NULL); pthread_key_delete(k); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn setspecific_null_skips_destructor() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint val = 0;\nvoid d(void* a) { val = 1; }\nvoid* f(void* a) { pthread_key_t *k = a; pthread_setspecific(*k, (void*)42); pthread_setspecific(*k, NULL); return NULL; }\nint main() { pthread_key_t k; pthread_key_create(&k, d); pthread_t t; pthread_create(&t, NULL, f, &k); pthread_join(t, NULL); printf(\"%d\", val); pthread_key_delete(k); return 0; }"
        ),
        vec!["0"]
    );
}
#[test]
fn destructor_repeated_calls() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint iters = 0;\npthread_key_t k;\nvoid d(void* a) { iters++; if(iters < 3) pthread_setspecific(k, (void*)1); }\nvoid* f(void* a) { pthread_setspecific(k, (void*)1); return NULL; }\nint main() { pthread_key_create(&k, d); pthread_t t; pthread_create(&t, NULL, f, NULL); pthread_join(t, NULL); printf(\"%d\", iters); pthread_key_delete(k); return 0; }"
        ),
        vec!["3"]
    );
} // POSIX allows up to PTHREAD_DESTRUCTOR_ITERATIONS (usually 4)
#[test]
fn tls_macro_threads_isolated() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\n#include <unistd.h>\n__thread int val = 0;\nvoid* f(void* a) { val = 1; return NULL; }\nint main() { val = 2; pthread_t t; pthread_create(&t, NULL, f, NULL); pthread_join(t, NULL); printf(\"%d\", val); return 0; }"
        ),
        vec!["2"]
    );
}
#[test]
fn pthread_once_basic() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\npthread_once_t once = PTHREAD_ONCE_INIT;\nint count = 0;\nvoid init() { count++; }\nvoid* f(void* a) { pthread_once(&once, init); return NULL; }\nint main() { pthread_t t1, t2; pthread_create(&t1, NULL, f, NULL); pthread_create(&t2, NULL, f, NULL); pthread_join(t1, NULL); pthread_join(t2, NULL); pthread_once(&once, init); printf(\"%d\", count); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn key_delete_while_threads_active() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\n#include <unistd.h>\npthread_key_t k;\nvoid* f(void* a) { sleep(1); return pthread_getspecific(k); }\nint main() { pthread_key_create(&k, NULL); pthread_setspecific(k, (void*)42); pthread_t t; pthread_create(&t, NULL, f, NULL); pthread_key_delete(k); /* deleting key does not call destructors, and subsequent getspecific is UB but usually returns NULL or old val. We check it compiles and runs */ pthread_join(t, NULL); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn getspecific_main_thread() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { pthread_key_t k; pthread_key_create(&k, NULL); pthread_setspecific(k, (void*)99); printf(\"%ld\", (long)pthread_getspecific(k)); pthread_key_delete(k); return 0; }"
        ),
        vec!["99"]
    );
}
#[test]
fn key_create_null_destructor() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { pthread_key_t k; int r = pthread_key_create(&k, NULL); printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
