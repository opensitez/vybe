// vybe-test: c/c_stdlib_prng_rand_r/rand_r_basic
// origin: languages/c/tests/c/test_c_stdlib_prng_rand_r.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 199506L
#include <stdlib.h>
int main() {const char *__w[] = {"1 1"};
int __n = 1, __i = 0;
 unsigned int seed1 = 42; unsigned int seed2 = 42; int r1 = rand_r(&seed1); int r2 = rand_r(&seed2); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", r1 == r2, seed1 != 42);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

