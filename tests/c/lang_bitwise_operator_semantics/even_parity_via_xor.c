// vybe-test: c/lang_bitwise_operator_semantics/even_parity_via_xor
// origin: languages/c/tests/c/test_lang_bitwise_operator_semantics.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
unsigned u=0b1110; int parity=u^(u>>1)^(u>>2)^(u>>3); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", parity & 1);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

