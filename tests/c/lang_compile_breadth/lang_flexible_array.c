// vybe-test: c/lang_compile_breadth/lang_flexible_array
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdlib.h>
struct B { int n; char d[]; };
int main() {
struct B *b=malloc(sizeof(*b)+2); free(b); return 0;
}

