// vybe-test: c/c_complex_arithmetic/complex_exp
// origin: languages/c/tests/c/test_c_complex_arithmetic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <complex.h>
#include <math.h>
int main() {const char *__w[] = {"-1.0"};
int __n = 1, __i = 0;
 double complex z = cexp(0.0 + M_PI * I); { char __t[512]; snprintf(__t, sizeof(__t), "%.1f", creal(z));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

