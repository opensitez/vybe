// vybe-test: c/stdlib_conversion_and_exit/atexit_two_handlers_compile
// origin: languages/c/tests/c/test_stdlib_conversion_and_exit.rs
// vybe-test-mode: compile
#include <stdlib.h>
void a(void){} void b(void){}
int main() {
atexit(a); atexit(b); return 0;
}

