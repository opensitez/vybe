// vybe-test: c/cover_threads_fenv/fegetround
// origin: languages/c/tests/c/test_cover_threads_fenv.rs
// vybe-test-mode: compile
#include <fenv.h>
int main() {
return fegetround();
}

