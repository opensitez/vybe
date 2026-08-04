// vybe-test: c/cover_breadth_batch_b/longjmp_compile
// origin: languages/c/tests/c/test_cover_breadth_batch_b.rs
// vybe-test-mode: compile
#include <setjmp.h>
int main() {
jmp_buf b; if(!setjmp(b)) longjmp(b,1); return 0;
}

