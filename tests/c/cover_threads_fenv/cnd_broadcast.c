// vybe-test: c/cover_threads_fenv/cnd_broadcast
// origin: languages/c/tests/c/test_cover_threads_fenv.rs
// vybe-test-mode: compile
#include <threads.h>
int main() {
cnd_t c; cnd_init(&c); cnd_broadcast(&c); cnd_destroy(&c); return 0;
}

