// vybe-test: c/c_logical_short_circuit_side_effects/short_circuit_or_basic
// origin: languages/c/tests/c/test_c_logical_short_circuit_side_effects.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"0"};
int __n = 1, __i = 0;
 int x=0; 1 || (x=1); { char __t[512]; snprintf(__t, sizeof(__t), "%d", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

