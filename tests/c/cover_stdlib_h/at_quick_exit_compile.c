// vybe-test: c/cover_stdlib_h/at_quick_exit_compile
// origin: languages/c/tests/c/test_cover_stdlib_h.rs
// vybe-test-mode: compile
#include <stdlib.h>
void h(void) {}
int main() {
at_quick_exit(h); return 0;
}

