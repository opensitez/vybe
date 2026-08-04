// vybe-test: c/cover_breadth_batch_a/stdlib_mkstemp
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <stdlib.h>
int main() {
char t[]="/tmp/vybeXXXXXX"; mkstemp(t); return 0;
}

