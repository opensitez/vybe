// vybe-test: c/lang_compile_breadth/lang_struct_nested_anon
// origin: languages/c/tests/c/test_lang_compile_breadth.rs
// vybe-test-mode: compile
#include <stdio.h>
struct S { struct { int x; } inner; };
int main() {
struct S s={.inner={1}}; return s.inner.x;
}

