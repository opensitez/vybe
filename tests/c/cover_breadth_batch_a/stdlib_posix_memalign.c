// vybe-test: c/cover_breadth_batch_a/stdlib_posix_memalign
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <stdlib.h>
int main() {
void *p=0; posix_memalign(&p,16,32); free(p); return 0;
}

