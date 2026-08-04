// vybe-test: c/cover_threads_fenv/feclearexcept
// origin: languages/c/tests/c/test_cover_threads_fenv.rs
// vybe-test-mode: compile
#include <fenv.h>
int main() {
feclearexcept(FE_ALL_EXCEPT); return 0;
}

