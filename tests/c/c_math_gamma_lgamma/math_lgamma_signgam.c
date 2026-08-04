// vybe-test: c/c_math_gamma_lgamma/math_lgamma_signgam
// origin: languages/c/tests/c/test_c_math_gamma_lgamma.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <math.h>
extern int signgam;
int main() {const char *__w[] = {"-1"};
int __n = 1, __i = 0;
 lgamma(-0.5); { char __t[512]; snprintf(__t, sizeof(__t), "%d", signgam);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

