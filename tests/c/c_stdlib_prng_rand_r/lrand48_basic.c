// vybe-test: c/c_stdlib_prng_rand_r/lrand48_basic
// origin: languages/c/tests/c/test_c_stdlib_prng_rand_r.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _XOPEN_SOURCE
#include <stdlib.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 srand48(42); long r1 = lrand48(); srand48(42); long r2 = lrand48(); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r1 == r2 && r1 >= 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

