// vybe-test: c/lang_preprocessor_breadth/defined_operator
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#if defined(__STDC__)
int g=1;
#endif
int main() {
return g;
}

