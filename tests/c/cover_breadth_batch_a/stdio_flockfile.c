// vybe-test: c/cover_breadth_batch_a/stdio_flockfile
// origin: languages/c/tests/c/test_cover_breadth_batch_a.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
flockfile(stdout); funlockfile(stdout); return 0;
}

