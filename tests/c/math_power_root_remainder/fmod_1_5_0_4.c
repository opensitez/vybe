// vybe-test: c/math_power_root_remainder/fmod_1_5_0_4
// origin: languages/c/tests/c/test_math_power_root_remainder.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <math.h>
int main() {
const char *__w[] = {"0.3\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%.1f\n", fmod(1.5, 0.4));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

