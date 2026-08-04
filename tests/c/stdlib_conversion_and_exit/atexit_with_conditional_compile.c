// vybe-test: c/stdlib_conversion_and_exit/atexit_with_conditional_compile
// origin: languages/c/tests/c/test_stdlib_conversion_and_exit.rs
// vybe-test-mode: compile
#include <stdlib.h>
void c(void){}
int main() {
if (1) atexit(c); return 0;
}

