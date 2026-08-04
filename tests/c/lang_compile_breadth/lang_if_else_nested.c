// vybe-test: c/lang_compile_breadth/lang_if_else_nested
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
int main() {
if(1) if(0) return 1; else return 2; return 0;
}

