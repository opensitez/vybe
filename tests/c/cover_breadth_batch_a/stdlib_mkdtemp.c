// vybe-test: c/cover_breadth_batch_a/stdlib_mkdtemp
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <stdlib.h>
int main() {
char t[]="/tmp/vybeXXXXXX"; mkdtemp(t); return 0;
}

