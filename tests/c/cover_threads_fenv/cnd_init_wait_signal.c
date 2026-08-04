// vybe-test: c/cover_threads_fenv/cnd_init_wait_signal
// origin: languages/c/tests/c/test_cover_threads_fenv.rs
// vybe-test-mode: compile
#include <threads.h>
int main() {
cnd_t c; mtx_t m; cnd_init(&c); mtx_init(&m, mtx_plain); cnd_signal(&c); cnd_destroy(&c); mtx_destroy(&m); return 0;
}

