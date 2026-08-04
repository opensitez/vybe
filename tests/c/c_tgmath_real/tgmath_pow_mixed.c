// vybe-test: c/c_tgmath_real/tgmath_pow_mixed
// origin: languages/c/tests/c/test_c_tgmath_real.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <tgmath.h>
int main() {const char *__w[] = {"8.0"};
int __n = 1, __i = 0;
 double x = 2.0; float y = 3.0f; { char __t[512]; snprintf(__t, sizeof(__t), "%.1f", pow(x, y));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

