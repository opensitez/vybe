// vybe-test: c/cover_threads_fenv/thrd_current
// origin: languages/c/tests/c/test_cover_threads_fenv.rs
// vybe-test-mode: compile
#include <threads.h>
int main() {
thrd_t t = thrd_current(); return t ? 0 : 0;
}

