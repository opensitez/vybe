// vybe-test: c/lang_preprocessor_breadth/if_expression
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#if 1+1==2
int c=3;
#endif
int main() {
return c;
}

