// vybe-test: c/lang_cast_expression_semantics/function_pointer_to_void_pointer
// origin: languages/c/tests/c/test_lang_cast_expression_semantics.rs
// vybe-test-mode: compile
#include <stdio.h>
int id(int x) { return x; }
int main() {
int (*fp)(int) = id; void *vp = (void *)fp; return vp != 0;
}

