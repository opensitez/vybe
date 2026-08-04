// vybe-test: c/cover_breadth_batch_a/stdlib_getenv_setenv
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <stdlib.h>
int main() {
setenv("VYBE_TEST","1",1); return getenv("VYBE_TEST") != 0;
}

