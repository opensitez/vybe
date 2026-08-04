// vybe-test: c/stdlib_conversion_and_exit/atexit_single_handler_compile
// origin: languages/c/tests/c/test_stdlib_conversion_and_exit.rs
// vybe-test-mode: compile
#include <stdlib.h>
void h(void){}
int main() {
atexit(h); return 0;
}

