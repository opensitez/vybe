// vybe-test: c/lang_arithmetic_promotion/unsigned_int_minus_char
// origin: languages/c/tests/c/test_lang_arithmetic_promotion.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"42\n"};
int __n = 1, __i = 0;
unsigned u=50u; char c=8; { char __t[512]; snprintf(__t, sizeof(__t), "%u\n", u-c);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

