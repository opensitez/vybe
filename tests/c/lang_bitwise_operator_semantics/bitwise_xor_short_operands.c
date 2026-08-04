// vybe-test: c/lang_bitwise_operator_semantics/bitwise_xor_short_operands
// origin: languages/c/tests/c/test_lang_bitwise_operator_semantics.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
short a=0x0C, b=0x0A; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a ^ b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

