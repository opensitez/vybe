// vybe-test: c/cover_breadth_batch_b/setjmp_buffer
// origin: languages/c/tests/c/test_cover_breadth_batch_b.rs
// vybe-test-mode: compile
#include <setjmp.h>
int main() {
jmp_buf b; return setjmp(b);
}

