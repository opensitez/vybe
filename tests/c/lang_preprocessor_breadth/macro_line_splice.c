// vybe-test: c/lang_preprocessor_breadth/macro_line_splice
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#define M 1 \
+ 2
int main() {
return M;
}

