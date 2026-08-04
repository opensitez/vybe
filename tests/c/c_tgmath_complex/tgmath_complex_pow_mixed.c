// vybe-test: c/c_tgmath_complex/tgmath_complex_pow_mixed
// origin: languages/c/tests/c/test_c_tgmath_complex.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <tgmath.h>
int main() {const char *__w[] = {"8.0"};
int __n = 1, __i = 0;
 double complex z = 2.0 + 0.0 * I; double x = 3.0; { char __t[512]; snprintf(__t, sizeof(__t), "%.1f", creal(pow(z, x)));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

