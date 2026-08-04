// vybe-test: c/limits_float_named_constants/named_flt_max_exceeds_one
// origin: languages/c/tests/c/test_limits_float_named_constants.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <float.h>
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", FLT_MAX > 1.0f);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

