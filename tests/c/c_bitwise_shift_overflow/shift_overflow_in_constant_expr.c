// vybe-test: c/c_bitwise_shift_overflow/shift_overflow_in_constant_expr
// origin: languages/c/tests/c/test_c_bitwise_shift_overflow.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"2147483648"};
int __n = 1, __i = 0;
 enum { A = 1U << 31 }; { char __t[512]; snprintf(__t, sizeof(__t), "%u", (unsigned int)A);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

