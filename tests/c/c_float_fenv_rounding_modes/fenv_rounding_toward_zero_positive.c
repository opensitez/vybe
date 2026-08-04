// vybe-test: c/c_float_fenv_rounding_modes/fenv_rounding_toward_zero_positive
// origin: languages/c/tests/c/test_c_float_fenv_rounding_modes.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <fenv.h>
#pragma STDC FENV_ACCESS ON
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 fesetround(FE_TOWARDZERO); double z = 1.0 / 3.0; { char __t[512]; snprintf(__t, sizeof(__t), "%d", z < 0.3333333333333334);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

