// vybe-test: c/cover_complex_limits/atomic_flag_clear
// origin: languages/c/tests/c/test_cover_complex_limits.rs
// vybe-test-mode: compile
#include <stdatomic.h>
int main() {
atomic_flag f=ATOMIC_FLAG_INIT; atomic_flag_clear(&f); return 0;
}

