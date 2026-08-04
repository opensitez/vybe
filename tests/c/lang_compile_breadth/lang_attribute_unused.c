// vybe-test: c/lang_compile_breadth/lang_attribute_unused
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
__attribute__((unused)) static int u;
int main() {
return 0;
}

