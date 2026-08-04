// vybe-test: c/stdlib_conversion_and_exit/atexit_multiple_same_handler_compile
// origin: languages/c/tests/c/test_stdlib_conversion_and_exit.rs
// vybe-test-mode: compile
#include <stdlib.h>
void same(void){}
int main() {
atexit(same); atexit(same); return 0;
}

