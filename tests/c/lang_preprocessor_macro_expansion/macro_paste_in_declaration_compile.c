// vybe-test: c/lang_preprocessor_macro_expansion/macro_paste_in_declaration_compile
// origin: languages/c/tests/c/test_lang_preprocessor_macro_expansion.rs
// vybe-test-mode: compile
#include <stdio.h>
#define TYPEDEF_NAME counter
int TYPEDEF_NAME = 0;
int main() {
return TYPEDEF_NAME;
}

