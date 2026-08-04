// vybe-test: c/lang_preprocessor_breadth/define_function_macro
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#define SQ(x) ((x)*(x))
int main() {
return SQ(3);
}

