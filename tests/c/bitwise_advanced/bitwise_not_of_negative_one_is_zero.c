// vybe-test: c/bitwise_advanced/bitwise_not_of_negative_one_is_zero
// origin: languages/c/tests/c/test_bitwise_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"0\n"};
int __n = 1, __i = 0;
int x = -1; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", ~x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

