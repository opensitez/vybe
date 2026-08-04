// vybe-test: c/lang_compile_breadth/lang_generic_default
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
#define T(x) _Generic((x), int:1, default:0)
int main() {
return T(1.0);
}

