// vybe-test: c/cover_breadth_batch_a/stdlib_putenv
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <stdlib.h>
int main() {
putenv("VYBE_TEST=2"); return 0;
}

