// vybe-test: c/c_comma_operator_precedence/comma_operator_void_type
// origin: languages/c/tests/c/test_c_comma_operator_precedence.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void f() {} int main() {const char *__w[] = {"5"};
int __n = 1, __i = 0;
 int a = (f(), 5); { char __t[512]; snprintf(__t, sizeof(__t), "%d", a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

