// vybe-test: c/c_float_hex_literals/float_hex_no_fraction
// origin: languages/c/tests/c/test_c_float_hex_literals.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"16.000000"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "%f", 0x2p3);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

