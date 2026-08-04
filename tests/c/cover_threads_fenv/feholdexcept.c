// vybe-test: c/cover_threads_fenv/feholdexcept
// origin: languages/c/tests/c/test_cover_threads_fenv.rs
// vybe-test-mode: compile
#include <fenv.h>
int main() {
fenv_t env; feholdexcept(&env); return 0;
}

