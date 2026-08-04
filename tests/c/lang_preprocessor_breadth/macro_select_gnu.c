// vybe-test: c/lang_preprocessor_breadth/macro_select_gnu
// origin: languages/c/tests/c/test_lang_preprocessor_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#define VAL(x) _Generic((x), int: 1, default: 0)
int main() {
return VAL(0);
}

