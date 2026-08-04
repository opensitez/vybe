// vybe-test: c/lang_preprocessor_breadth/undef_macro
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#define Z 1
#undef Z
int f=2;
int main() {
return f;
}

