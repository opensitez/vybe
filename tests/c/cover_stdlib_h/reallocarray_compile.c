// vybe-test: c/cover_stdlib_h/reallocarray_compile
// origin: languages/c/tests/c/test_cover_stdlib_h.rs
// vybe-test-mode: compile
#include <stdlib.h>
int main() {
void *p = reallocarray(0, 2, 4); free(p); return 0;
}

