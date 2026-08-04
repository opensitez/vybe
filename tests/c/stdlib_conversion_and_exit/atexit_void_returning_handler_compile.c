// vybe-test: c/stdlib_conversion_and_exit/atexit_void_returning_handler_compile
// origin: languages/c/tests/c/test_stdlib_conversion_and_exit.rs
// vybe-test-mode: compile
#include <stdlib.h>
void vr(void){ return; }
int main() {
atexit(vr); return 0;
}

