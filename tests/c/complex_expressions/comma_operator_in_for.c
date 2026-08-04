// vybe-test: c/complex_expressions/comma_operator_in_for
// origin: languages/c/tests/c/test_complex_expressions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"1 10\n", "2 9\n"};
int __n = 2, __i = 0;
int a=0, b=0;
for (a=1, b=10; a < 3; a++, b--) { { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", a, b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

