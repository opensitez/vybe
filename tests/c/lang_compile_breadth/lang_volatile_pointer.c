// vybe-test: c/lang_compile_breadth/lang_volatile_pointer
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
volatile int v=1; volatile int *p=&v; return *p;
}

