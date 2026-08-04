// vybe-test: c/math_extended/isnan_detects_nan
// origin: languages/c/tests/c/test_math_extended.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
double nan_val = 0.0 / 0.0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", isnan(nan_val) ? 1 : 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

