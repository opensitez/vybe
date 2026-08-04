// vybe-test: c/lang_functions_types/const_param_pointer
// origin: languages/c/tests/c/test_lang_functions_types.rs
// vybe-test-mode: compile
#include <stdio.h>
void ro(const int *p){}
int main() {
int x=1; ro(&x); return 0;
}

