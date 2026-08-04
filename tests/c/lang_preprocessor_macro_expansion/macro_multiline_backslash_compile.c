// vybe-test: c/lang_preprocessor_macro_expansion/macro_multiline_backslash_compile
// origin: languages/c/tests/c/test_lang_preprocessor_macro_expansion.rs
// vybe-test-mode: compile
#include <stdio.h>
#define INC(x) \
((x)+1)
int main() {
return INC(1);
}

