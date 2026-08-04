// vybe-test: c/c_ternary_nested/ternary_nested_different_types
// origin: languages/c/tests/c/test_c_ternary_nested.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 double a = 1 ? 0 ? 1 : 2.5 : 3; { char __t[512]; snprintf(__t, sizeof(__t), "%d", a > 2.0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

