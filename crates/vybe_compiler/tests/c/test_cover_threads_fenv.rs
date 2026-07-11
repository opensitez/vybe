//! C11 threads.h and fenv.h — one API per compile smoke.

c_compile_cases! {
    thrd_create_join => {
        includes: ["<stdio.h>", "<threads.h>"],
        decls: "int worker(void *arg) { (void)arg; return 0; }",
        body: "thrd_t t; thrd_create(&t, worker, 0); thrd_join(t, 0); return 0;"
    },
    thrd_yield => { includes: ["<threads.h>"], decls: "", body: "thrd_yield(); return 0;" },
    thrd_sleep => { includes: ["<threads.h>"], decls: "", body: "struct timespec ts = {0, 0}; thrd_sleep(&ts, 0); return 0;" },
    thrd_current => { includes: ["<threads.h>"], decls: "", body: "thrd_t t = thrd_current(); return t ? 0 : 0;" },
    thrd_equal => { includes: ["<threads.h>"], decls: "", body: "return thrd_equal(thrd_current(), thrd_current());" },
    mtx_init_lock_unlock => { includes: ["<threads.h>"], decls: "", body: "mtx_t m; mtx_init(&m, mtx_plain); mtx_lock(&m); mtx_unlock(&m); mtx_destroy(&m); return 0;" },
    mtx_trylock => { includes: ["<threads.h>"], decls: "", body: "mtx_t m; mtx_init(&m, mtx_plain); mtx_trylock(&m); mtx_unlock(&m); mtx_destroy(&m); return 0;" },
    cnd_init_wait_signal => {
        includes: ["<threads.h>"],
        decls: "",
        body: "cnd_t c; mtx_t m; cnd_init(&c); mtx_init(&m, mtx_plain); cnd_signal(&c); cnd_destroy(&c); mtx_destroy(&m); return 0;"
    },
    cnd_broadcast => { includes: ["<threads.h>"], decls: "", body: "cnd_t c; cnd_init(&c); cnd_broadcast(&c); cnd_destroy(&c); return 0;" },
    tss_create_set_get => { includes: ["<threads.h>"], decls: "", body: "tss_t key; tss_create(&key, 0); tss_set(key, (void*)1); tss_get(key); tss_delete(key); return 0;" },
    call_once => {
        includes: ["<threads.h>"],
        decls: "void once_fn(void) {}",
        body: "once_flag f = ONCE_FLAG_INIT; call_once(&f, once_fn); return 0;"
    },
    feclearexcept => { includes: ["<fenv.h>"], decls: "", body: "feclearexcept(FE_ALL_EXCEPT); return 0;" },
    fetestexcept => { includes: ["<fenv.h>"], decls: "", body: "return fetestexcept(FE_DIVBYZERO);" },
    feholdexcept => { includes: ["<fenv.h>"], decls: "", body: "fenv_t env; feholdexcept(&env); return 0;" },
    fesetround => { includes: ["<fenv.h>"], decls: "", body: "fesetround(FE_TONEAREST); return 0;" },
    fegetround => { includes: ["<fenv.h>"], decls: "", body: "return fegetround();" },
}
