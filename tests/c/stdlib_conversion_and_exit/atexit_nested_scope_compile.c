// vybe-test: c/stdlib_conversion_and_exit/atexit_nested_scope_compile
// origin: languages/c/tests/c/test_stdlib_conversion_and_exit.rs
// vybe-test-mode: compile
#include <stdlib.h>
void outer(void){}
int main() {
{ atexit(outer); } return 0;
}

