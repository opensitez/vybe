// vybe-test: c/cover_threads_fenv/thrd_equal
// origin: languages/c/tests/c/test_cover_threads_fenv.rs
// vybe-test-mode: compile
#include <threads.h>
int main() {
return thrd_equal(thrd_current(), thrd_current());
}

