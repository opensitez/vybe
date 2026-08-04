// vybe-test: c/cover_threads_fenv/thrd_sleep
// origin: languages/c/tests/c/test_cover_threads_fenv.rs
// vybe-test-mode: compile
#include <threads.h>
int main() {
struct timespec ts = {0, 0}; thrd_sleep(&ts, 0); return 0;
}

