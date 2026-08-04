// vybe-test: c/cover_breadth_batch_a/string_bzero
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <strings.h>
int main() {
char b[4]; bzero(b,4); return 0;
}

