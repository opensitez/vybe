// vybe-test: c/cover_stdlib_h/aligned_alloc_compile
// origin: languages/c/tests/c/test_cover_stdlib_h.rs
// vybe-test-mode: compile
#include <stdlib.h>
int main() {
void *p = aligned_alloc(16, 32); free(p); return 0;
}

