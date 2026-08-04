// vybe-test: c/c_bitwise_shift_overflow/shift_count_exceeds_promotion
// origin: languages/c/tests/c/test_c_bitwise_shift_overflow.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"65536"};
int __n = 1, __i = 0;
 char c = 1; /* c << 16 is fine if int is 32-bit because c promotes to int before shift */ { char __t[512]; snprintf(__t, sizeof(__t), "%d", c << 16);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

