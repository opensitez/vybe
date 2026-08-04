// vybe-test: c/c_math_nextafter_copysign/math_nextafter_zero_to_pos
// origin: languages/c/tests/c/test_c_math_nextafter_copysign.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <math.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 double n = nextafter(0.0, 1.0); { char __t[512]; snprintf(__t, sizeof(__t), "%d", n > 0.0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

