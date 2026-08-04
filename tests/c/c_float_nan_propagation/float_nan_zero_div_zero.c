// vybe-test: c/c_float_nan_propagation/float_nan_zero_div_zero
// origin: languages/c/tests/c/test_c_float_nan_propagation.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <math.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 double n = 0.0 / 0.0; { char __t[512]; snprintf(__t, sizeof(__t), "%d", isnan(n));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

