// vybe-test: c/lang_compile_breadth/lang_empty_struct
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
struct E {};
int main() {
struct E e; return sizeof(e)>=1;
}

