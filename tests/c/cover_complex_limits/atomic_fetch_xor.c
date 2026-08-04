// vybe-test: c/cover_complex_limits/atomic_fetch_xor
// origin: languages/c/tests/c/test_cover_complex_limits.rs
// vybe-test-mode: compile
#include <stdatomic.h>
int main() {
atomic_int x=3; atomic_fetch_xor(&x,1); return atomic_load(&x);
}

