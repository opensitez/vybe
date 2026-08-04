// vybe-test: c/cover_breadth_batch_a/stdio_scanf_compile
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
int x; sscanf("42","%d",&x); return x;
}

