// vybe-test: c/c_comma_operator_precedence/comma_operator_in_for_init
// origin: languages/c/tests/c/test_c_comma_operator_precedence.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 int a, b; for (a=1, b=2; a<2; a++) { char __t[512]; snprintf(__t, sizeof(__t), "%d", b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

