// vybe-test: c/stdlib_conversion_and_exit/atexit_with_static_function_compile
// origin: languages/c/tests/c/test_stdlib_conversion_and_exit.rs
// vybe-test-mode: compile
#include <stdlib.h>
static void s(void){}
int main() {
atexit(s); return 0;
}

