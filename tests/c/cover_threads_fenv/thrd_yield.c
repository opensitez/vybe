// vybe-test: c/cover_threads_fenv/thrd_yield
// origin: languages/c/tests/c/test_cover_threads_fenv.rs
// vybe-test-mode: compile
#include <threads.h>
int main() {
thrd_yield(); return 0;
}

