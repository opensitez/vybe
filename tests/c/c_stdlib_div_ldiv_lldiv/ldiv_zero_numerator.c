// vybe-test: c/c_stdlib_div_ldiv_lldiv/ldiv_zero_numerator
// origin: languages/c/tests/c/test_c_stdlib_div_ldiv_lldiv.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int main() {const char *__w[] = {"0 0"};
int __n = 1, __i = 0;
 ldiv_t d = ldiv(0L, 5L); { char __t[512]; snprintf(__t, sizeof(__t), "%ld %ld", d.quot, d.rem);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

