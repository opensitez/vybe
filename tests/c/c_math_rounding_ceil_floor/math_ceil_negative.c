// vybe-test: c/c_math_rounding_ceil_floor/math_ceil_negative
// origin: languages/c/tests/c/test_c_math_rounding_ceil_floor.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <math.h>
int main() {const char *__w[] = {"-2.0"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "%.1f", ceil(-2.3));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

