// vybe-test: c/lang_functions_types/attribute_unused_compile
// origin: languages/c/tests/c/test_lang_functions_types.rs
// vybe-test-mode: compile
#include <stdio.h>
__attribute__((unused)) static int u = 0;
int main() {
return 0;
}

