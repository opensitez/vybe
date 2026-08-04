// vybe-test: c/lang_compile_breadth/lang_union_in_struct
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
struct S { union { int i; char c; } u; };
int main() {
struct S s; s.u.i=1; return s.u.i;
}

