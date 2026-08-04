// vybe-test: c/complex_numbers/float_complex_type
// origin: languages/c/tests/c/test_complex_numbers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

#include <stdio.h>
#include <complex.h>
int main() {const char *__w[] = {"2.0\n"};
int __n = 1, __i = 0;

    float complex z = 2.0f + 3.0f * I;
    { char __t[512]; snprintf(__t, sizeof(__t), "%.1f\n", crealf(z));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
    if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

