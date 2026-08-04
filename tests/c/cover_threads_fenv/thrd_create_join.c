// vybe-test: c/cover_threads_fenv/thrd_create_join
// origin: languages/c/tests/c/test_cover_threads_fenv.rs
// vybe-test-mode: compile
#include <stdio.h>
#include <threads.h>
int worker(void *arg) { (void)arg; return 0; }
int main() {
thrd_t t; thrd_create(&t, worker, 0); thrd_join(t, 0); return 0;
}

