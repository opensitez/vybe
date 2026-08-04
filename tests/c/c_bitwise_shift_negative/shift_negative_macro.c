// vybe-test: c/c_bitwise_shift_negative/shift_negative_macro
// origin: languages/c/tests/c/test_c_bitwise_shift_negative.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define SHIFT(x) ((x) >> 1)
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int x = -2; { char __t[512]; snprintf(__t, sizeof(__t), "%d", SHIFT(x) == -1 || SHIFT(x) > 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

