// vybe-test: c/cover_threads_fenv/call_once
// origin: languages/c/tests/c/test_cover_threads_fenv.rs
// vybe-test-mode: compile
#include <threads.h>
void once_fn(void) {}
int main() {
once_flag f = ONCE_FLAG_INIT; call_once(&f, once_fn); return 0;
}

