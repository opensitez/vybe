// vybe-test: c/lang_vla_stack_arrays/vla_multidim_parameter_compile
// origin: languages/c/tests/c/test_lang_vla_stack_arrays.rs
// vybe-test-mode: compile
#include <stdio.h>
void g(int r, int c, int m[r][c]); void g(int r, int c, int m[r][c]){}
int main() {
return 0;
}

