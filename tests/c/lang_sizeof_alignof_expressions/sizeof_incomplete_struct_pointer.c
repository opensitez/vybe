// vybe-test: c/lang_sizeof_alignof_expressions/sizeof_incomplete_struct_pointer
// origin: languages/c/tests/c/test_lang_sizeof_alignof_expressions.rs
// vybe-test-mode: compile
#include <stdio.h>
struct Incomplete; struct Incomplete *p;
int main() {
return (int)sizeof(p);
}

