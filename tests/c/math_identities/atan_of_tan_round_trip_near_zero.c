// vybe-test: c/math_identities/atan_of_tan_round_trip_near_zero
// origin: languages/c/tests/c/test_math_identities.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"0.200\n"};
int __n = 1, __i = 0;
double x = 0.2; { char __t[512]; snprintf(__t, sizeof(__t), "%.3f\n", atan(tan(x)));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

