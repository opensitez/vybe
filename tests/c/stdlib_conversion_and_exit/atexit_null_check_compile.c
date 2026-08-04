// vybe-test: c/stdlib_conversion_and_exit/atexit_null_check_compile
// origin: languages/c/tests/c/test_stdlib_conversion_and_exit.rs
// vybe-test-mode: compile
#include <stdlib.h>
void z(void){}
int main() {
return atexit(z) == 0;
}

