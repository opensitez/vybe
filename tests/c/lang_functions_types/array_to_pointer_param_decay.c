// vybe-test: c/lang_functions_types/array_to_pointer_param_decay
// origin: languages/c/tests/c/test_lang_functions_types.rs
// vybe-test-mode: compile
#include <stdio.h>
void take(int *a){}
int main() {
int x[2]={1,2}; take(x); return 0;
}

