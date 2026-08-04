// vybe-test: c/c_typeof_gnu_extension/typeof_expr_unevaluated
// origin: languages/c/tests/c/test_c_typeof_gnu_extension.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int f(int *x) { (*x)++; return 1; } int main() {const char *__w[] = {"0"};
int __n = 1, __i = 0;
 int x=0; typeof(f(&x)) y = 5; { char __t[512]; snprintf(__t, sizeof(__t), "%d", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

