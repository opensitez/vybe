// vybe-test: c/c_stdlib_prng_rand_r/lcong48_basic
// origin: languages/c/tests/c/test_c_stdlib_prng_rand_r.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _XOPEN_SOURCE
#include <stdlib.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 unsigned short param[7] = {1,2,3,4,5,6,7}; lcong48(param); double r = drand48(); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r >= 0.0 && r < 1.0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

