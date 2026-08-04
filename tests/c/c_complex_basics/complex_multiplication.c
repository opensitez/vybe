// vybe-test: c/c_complex_basics/complex_multiplication
// origin: languages/c/tests/c/test_c_complex_basics.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <complex.h>
int main() {const char *__w[] = {"-4.0 7.0"};
int __n = 1, __i = 0;
 double complex z1 = 1.0 + 2.0 * I; double complex z2 = 2.0 + 3.0 * I; double complex z3 = z1 * z2; /* (1+2i)*(2+3i) = 2 + 3i + 4i - 6 = -4 + 7i */ { char __t[512]; snprintf(__t, sizeof(__t), "%.1f %.1f", creal(z3), cimag(z3));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

