// vybe-test: c/stdlib_conversion_and_exit/atexit_three_handlers_lifo_compile
// origin: languages/c/tests/c/test_stdlib_conversion_and_exit.rs
// vybe-test-mode: compile
#include <stdlib.h>
void f(void){} void g(void){} void h(void){}
int main() {
atexit(f); atexit(g); atexit(h); return 0;
}

