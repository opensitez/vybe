// vybe-test: c/stdlib_conversion_and_exit/atexit_before_return_compile
// origin: languages/c/tests/c/test_stdlib_conversion_and_exit.rs
// vybe-test-mode: compile
#include <stdlib.h>
void fin(void){}
int main() {
atexit(fin); return 0;
}

