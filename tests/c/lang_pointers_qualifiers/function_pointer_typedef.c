// vybe-test: c/lang_pointers_qualifiers/function_pointer_typedef
// origin: languages/c/tests/c/test_lang_pointers_qualifiers.rs
// vybe-test-mode: compile
#include <stdio.h>
typedef int (*op_t)(int,int);
int main() {
op_t f = 0; return 0;
}

