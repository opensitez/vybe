// vybe-test: c/c_sizeof_expressions_unevaluated/sizeof_expr_function_call
// origin: languages/c/tests/c/test_c_sizeof_expressions_unevaluated.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int f(int *x) { (*x)++; return 1; } int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int x=1; int s = sizeof(f(&x)); { char __t[512]; snprintf(__t, sizeof(__t), "%d", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

