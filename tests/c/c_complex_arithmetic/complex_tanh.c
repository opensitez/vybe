// vybe-test: c/c_complex_arithmetic/complex_tanh
// origin: languages/c/tests/c/test_c_complex_arithmetic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <complex.h>
int main() {const char *__w[] = {"1.55741"};
int __n = 1, __i = 0;
 double complex z = ctanh(0.0 + 1.0 * I); { char __t[512]; snprintf(__t, sizeof(__t), "%.5f", cimag(z));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

