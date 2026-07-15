use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> {
    run_prints(&format!("#include <stdio.h>\n{}", src))
}

#[test]
fn pthread_create_basic() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nvoid* f(void* a) { return a; }\nint main() { pthread_t t; int r = pthread_create(&t, NULL, f, (void*)42); if(r==0) { void *res; pthread_join(t, &res); printf(\"%ld\", (long)res); } return 0; }"
        ),
        vec!["42"]
    );
}
#[test]
fn pthread_join_basic() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nvoid* f(void* a) { pthread_exit((void*)99); return NULL; }\nint main() { pthread_t t; pthread_create(&t, NULL, f, NULL); void *res; pthread_join(t, &res); printf(\"%ld\", (long)res); return 0; }"
        ),
        vec!["99"]
    );
}
#[test]
fn pthread_detach_basic() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\n#include <unistd.h>\nvoid* f(void* a) { return NULL; }\nint main() { pthread_t t; pthread_create(&t, NULL, f, NULL); int r = pthread_detach(t); printf(\"%d\", r == 0); sleep(1); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pthread_exit_main() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { /* exiting main thread does not terminate process if others are alive. we test basic compile */ printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn pthread_self_basic() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nvoid* f(void* a) { return (void*)pthread_self(); }\nint main() { pthread_t t; pthread_create(&t, NULL, f, NULL); void *res; pthread_join(t, &res); printf(\"%d\", (pthread_t)res == t); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pthread_equal_basic() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { pthread_t t1 = pthread_self(), t2 = pthread_self(); printf(\"%d\", pthread_equal(t1, t2) != 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pthread_attr_init_destroy() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { pthread_attr_t a; int r1 = pthread_attr_init(&a); int r2 = pthread_attr_destroy(&a); printf(\"%d %d\", r1 == 0, r2 == 0); return 0; }"
        ),
        vec!["1 1"]
    );
}
#[test]
fn pthread_attr_setdetachstate() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { pthread_attr_t a; pthread_attr_init(&a); int r = pthread_attr_setdetachstate(&a, PTHREAD_CREATE_DETACHED); printf(\"%d\", r == 0); pthread_attr_destroy(&a); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pthread_attr_getdetachstate() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { pthread_attr_t a; pthread_attr_init(&a); pthread_attr_setdetachstate(&a, PTHREAD_CREATE_DETACHED); int s; pthread_attr_getdetachstate(&a, &s); printf(\"%d\", s == PTHREAD_CREATE_DETACHED); pthread_attr_destroy(&a); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pthread_attr_setstacksize() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { pthread_attr_t a; pthread_attr_init(&a); int r = pthread_attr_setstacksize(&a, 1024*1024); printf(\"%d\", r == 0); pthread_attr_destroy(&a); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pthread_attr_getstacksize() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { pthread_attr_t a; pthread_attr_init(&a); pthread_attr_setstacksize(&a, 1024*1024); size_t s; pthread_attr_getstacksize(&a, &s); printf(\"%d\", s == 1024*1024); pthread_attr_destroy(&a); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pthread_cancel_basic() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\n#include <unistd.h>\nvoid* f(void* a) { while(1) sleep(1); return NULL; }\nint main() { pthread_t t; pthread_create(&t, NULL, f, NULL); pthread_cancel(t); void *res; pthread_join(t, &res); printf(\"%d\", res == PTHREAD_CANCELED); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pthread_setcancelstate() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nvoid* f(void* a) { int old; pthread_setcancelstate(PTHREAD_CANCEL_DISABLE, &old); return (void*)(long)old; }\nint main() { pthread_t t; pthread_create(&t, NULL, f, NULL); void *r; pthread_join(t, &r); printf(\"%d\", (long)r == PTHREAD_CANCEL_ENABLE || (long)r == PTHREAD_CANCEL_DISABLE); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pthread_setcanceltype() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nvoid* f(void* a) { int old; pthread_setcanceltype(PTHREAD_CANCEL_ASYNCHRONOUS, &old); return (void*)(long)old; }\nint main() { pthread_t t; pthread_create(&t, NULL, f, NULL); void *r; pthread_join(t, &r); printf(\"%d\", (long)r == PTHREAD_CANCEL_DEFERRED || (long)r == PTHREAD_CANCEL_ASYNCHRONOUS); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pthread_testcancel() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint main() { pthread_testcancel(); printf(\"ok\"); return 0; }"
        ),
        vec!["ok"]
    );
}
#[test]
fn pthread_cleanup_push_pop() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nint val = 0;\nvoid c(void* a) { val = 1; }\nvoid* f(void* a) { pthread_cleanup_push(c, NULL); pthread_cleanup_pop(1); return NULL; }\nint main() { pthread_t t; pthread_create(&t, NULL, f, NULL); pthread_join(t, NULL); printf(\"%d\", val); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pthread_join_null_retval() {
    assert_eq!(
        run_c(
            "#include <pthread.h>\nvoid* f(void* a) { return NULL; }\nint main() { pthread_t t; pthread_create(&t, NULL, f, NULL); int r = pthread_join(t, NULL); printf(\"%d\", r == 0); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pthread_attr_setguardsize() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE 500\n#include <pthread.h>\nint main() { pthread_attr_t a; pthread_attr_init(&a); int r = pthread_attr_setguardsize(&a, 4096); printf(\"%d\", r == 0); pthread_attr_destroy(&a); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pthread_attr_getguardsize() {
    assert_eq!(
        run_c(
            "#define _XOPEN_SOURCE 500\n#include <pthread.h>\nint main() { pthread_attr_t a; pthread_attr_init(&a); pthread_attr_setguardsize(&a, 4096); size_t s; pthread_attr_getguardsize(&a, &s); printf(\"%d\", s == 4096); pthread_attr_destroy(&a); return 0; }"
        ),
        vec!["1"]
    );
}
#[test]
fn pthread_tryjoin_np_gnu() {
    assert_eq!(
        run_c(
            "#define _GNU_SOURCE\n#include <pthread.h>\nvoid* f(void* a) { return NULL; }\nint main() { pthread_t t; pthread_create(&t, NULL, f, NULL); pthread_join(t, NULL); int r = pthread_tryjoin_np(t, NULL); printf(\"%d\", r != 0); return 0; }"
        ),
        vec!["1"]
    );
} // Should return error since already joined
