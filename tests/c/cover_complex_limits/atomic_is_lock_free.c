// vybe-test: c/cover_complex_limits/atomic_is_lock_free
// origin: languages/c/tests/c/test_cover_complex_limits.rs
// vybe-test-mode: compile
#include <stdatomic.h>
int main() {
return atomic_is_lock_free(&(atomic_int){0});
}

