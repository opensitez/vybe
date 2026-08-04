// vybe-test: c/c_bitwise_shift_negative/shift_negative_in_unsigned_context
// origin: languages/c/tests/c/test_c_bitwise_shift_negative.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 unsigned int x = (unsigned int)-4 >> 1; { char __t[512]; snprintf(__t, sizeof(__t), "%d", x > 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

