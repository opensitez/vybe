// vybe-test: c/cover_breadth_batch_a/setjmp_sigsetjmp
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <setjmp.h>
int main() {
sigjmp_buf b; return sigsetjmp(b,1);
}

