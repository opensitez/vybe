// vybe-test: c/cover_complex_limits/atomic_thread_fence
// origin: languages/c/tests/c/test_cover_complex_limits.rs
// vybe-test-mode: compile
#include <stdatomic.h>
int main() {
atomic_thread_fence(memory_order_seq_cst); return 0;
}

