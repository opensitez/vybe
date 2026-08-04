// vybe-test: c/lang_compile_breadth/lang_zero_length_array_ext
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
struct Z { int n; int a[0]; };
int main() {
return sizeof(struct Z);
}

