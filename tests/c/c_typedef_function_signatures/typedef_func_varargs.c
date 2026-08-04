// vybe-test: c/c_typedef_function_signatures/typedef_func_varargs
// origin: languages/c/tests/c/test_c_typedef_function_signatures.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int (*F)(const char*, ...); #include <stdio.h>
int main() { F p = printf; p("%d", 42); return 0; }

