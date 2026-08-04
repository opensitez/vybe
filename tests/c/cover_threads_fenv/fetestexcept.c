// vybe-test: c/cover_threads_fenv/fetestexcept
// origin: languages/c/tests/c/test_cover_threads_fenv.rs
// vybe-test-mode: compile
#include <fenv.h>
int main() {
return fetestexcept(FE_DIVBYZERO);
}

