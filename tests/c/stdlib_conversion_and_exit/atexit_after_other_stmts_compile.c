// vybe-test: c/stdlib_conversion_and_exit/atexit_after_other_stmts_compile
// origin: languages/c/tests/c/test_stdlib_conversion_and_exit.rs
// vybe-test-mode: compile
#include <stdlib.h>
void done(void){}
int main() {
int x=1; (void)x; atexit(done); return 0;
}

