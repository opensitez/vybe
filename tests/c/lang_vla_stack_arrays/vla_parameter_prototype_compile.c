// vybe-test: c/lang_vla_stack_arrays/vla_parameter_prototype_compile
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
// vybe-test-mode: compile
#include <stdio.h>
void f(int n, int a[n]); void f(int n, int a[n]){}
int main() {
return 0;
}

