// vybe-test: c/math_rounding/ceil_then_subtract_fraction
// origin: languages/c/tests/c/test_math_rounding.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <math.h>
int main() {
const char *__w[] = {"8.75\n"};
int __n = 1, __i = 0;
double x = ceil(8.1) - 0.25;
{ char __t[512]; snprintf(__t, sizeof(__t), "%.2f\n", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

