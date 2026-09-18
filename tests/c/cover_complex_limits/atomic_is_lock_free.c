// vybe-test: c/cover_complex_limits/atomic_is_lock_free
// origin: languages/c/tests/c/test_cover_complex_limits.rs
// vybe-test-mode: compile
#include <stdatomic.h>
int main() {
atomic_int a = 0;
return atomic_is_lock_free(&a) ? 0 : 0;
}

