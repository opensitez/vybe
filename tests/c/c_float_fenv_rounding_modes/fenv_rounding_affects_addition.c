// vybe-test: c/c_float_fenv_rounding_modes/fenv_rounding_affects_addition
// origin: languages/c/tests/c/test_c_float_fenv_rounding_modes.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <fenv.h>
#include <stdio.h>
#pragma STDC FENV_ACCESS ON
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 fesetround(FE_UPWARD); double x = 1.0; double y = 3.0; double z = x / y; { char __t[512]; snprintf(__t, sizeof(__t), "%d", z > 0.3333333333333333);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

