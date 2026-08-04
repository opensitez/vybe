// vybe-test: c/cover_breadth_batch_a/stdlib_realpath
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <stdlib.h>
int main() {
char *p=realpath(".",0); if(p) free(p); return 0;
}

