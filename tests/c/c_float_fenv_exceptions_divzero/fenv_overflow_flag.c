// vybe-test: c/c_float_fenv_exceptions_divzero/fenv_overflow_flag
// origin: languages/c/tests/c/test_c_float_fenv_exceptions_divzero.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <fenv.h>
#include <float.h>
#pragma STDC FENV_ACCESS ON
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 feclearexcept(FE_ALL_EXCEPT); double x = DBL_MAX * 2.0; { char __t[512]; snprintf(__t, sizeof(__t), "%d", (fetestexcept(FE_OVERFLOW) != 0));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

