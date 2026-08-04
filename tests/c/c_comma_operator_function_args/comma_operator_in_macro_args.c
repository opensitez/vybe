// vybe-test: c/c_comma_operator_function_args/comma_operator_in_macro_args
// origin: languages/c/tests/c/test_c_comma_operator_function_args.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define M(a) printf("%d", a)
int main() { M((1, 2)); return 0; }

