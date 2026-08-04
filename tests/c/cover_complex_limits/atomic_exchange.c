// vybe-test: c/cover_complex_limits/atomic_exchange
// origin: languages/c/tests/c/test_cover_complex_limits.rs
// vybe-test-mode: compile
#include <stdatomic.h>
int main() {
atomic_int x=0; return atomic_exchange(&x,1);
}

