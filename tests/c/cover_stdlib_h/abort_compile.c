// vybe-test: c/cover_stdlib_h/abort_compile
// origin: languages/c/tests/c/test_cover_stdlib_h.rs
// vybe-test-mode: compile
#include <stdlib.h>
int main() {
if (0) abort(); return 0;
}

