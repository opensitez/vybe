// vybe-test: c/lang_bitwise_operator_semantics/bitwise_and_preserves_common_bits
// origin: languages/c/tests/c/test_lang_bitwise_operator_semantics.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"136\n"};
int __n = 1, __i = 0;
unsigned a=0b11001100, b=0b10101010; { char __t[512]; snprintf(__t, sizeof(__t), "%u\n", a & b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

