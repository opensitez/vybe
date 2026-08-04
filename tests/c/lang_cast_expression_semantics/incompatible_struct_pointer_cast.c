// vybe-test: c/lang_cast_expression_semantics/incompatible_struct_pointer_cast
// origin: languages/c/tests/c/test_lang_cast_expression_semantics.rs
// vybe-test-mode: compile
#include <stdio.h>
struct A { int x; }; struct B { int y; };
int main() {
struct A a = {1}; struct B *bp = (struct B *)&a; return bp->y;
}

