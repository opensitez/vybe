// vybe-test: c/lang_preprocessor_conditional_code/elif_nested_compile
// origin: languages/c/tests/c/test_lang_preprocessor_conditional_code.rs
// vybe-test-mode: compile
#include <stdio.h>
#if 0
int a=1;
#elif 1
int a=2;
#endif
int main() {
return a;
}

