// vybe-test: c/c_math_rounding_trunc_round/math_lrint_positive
// origin: languages/c/tests/c/test_c_math_rounding_trunc_round.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <math.h>
#include <fenv.h>
#pragma STDC FENV_ACCESS ON
int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 fesetround(FE_DOWNWARD); { char __t[512]; snprintf(__t, sizeof(__t), "%ld", lrint(2.9));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

