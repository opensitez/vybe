// vybe-test: c/cover_threads_fenv/mtx_init_lock_unlock
// origin: languages/c/tests/c/test_cover_threads_fenv.rs
// vybe-test-mode: compile
#include <threads.h>
int main() {
mtx_t m; mtx_init(&m, mtx_plain); mtx_lock(&m); mtx_unlock(&m); mtx_destroy(&m); return 0;
}

