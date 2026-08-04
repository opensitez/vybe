// vybe-test: c/c_stdlib_div_ldiv_lldiv/imaxdiv_basic
// origin: languages/c/tests/c/test_c_stdlib_div_ldiv_lldiv.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <inttypes.h>
int main() {const char *__w[] = {"3 1"};
int __n = 1, __i = 0;
 imaxdiv_t d = imaxdiv((intmax_t)10, (intmax_t)3); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", (int)d.quot, (int)d.rem);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

