// vybe-test: c/cover_complex_limits/atomic_fetch_or
// origin: languages/c/tests/c/test_cover_complex_limits.rs
// vybe-test-mode: compile
#include <stdatomic.h>
int main() {
atomic_int x=1; atomic_fetch_or(&x,2); return atomic_load(&x);
}

