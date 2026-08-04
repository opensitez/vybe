// vybe-test: c/operators_arithmetic/division_on_double_variables_does_not_truncate
// origin: languages/c/tests/c/test_operators_arithmetic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"3.5\n"};
int __n = 1, __i = 0;
double a = 7.0, b = 2.0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%.1f\n", a / b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

