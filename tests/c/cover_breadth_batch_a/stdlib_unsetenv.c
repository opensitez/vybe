// vybe-test: c/cover_breadth_batch_a/stdlib_unsetenv
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <stdlib.h>
int main() {
unsetenv("VYBE_TEST"); return 0;
}

