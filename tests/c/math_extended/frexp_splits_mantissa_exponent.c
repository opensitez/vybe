// vybe-test: c/math_extended/frexp_splits_mantissa_exponent
// origin: languages/c/tests/c/test_math_extended.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"0.5000 4\n"};
int __n = 1, __i = 0;
int e;
double m = frexp(8.0, &e);
{ char __t[512]; snprintf(__t, sizeof(__t), "%.4f %d\n", m, e);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

