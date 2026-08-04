// vybe-test: c/lang_arithmetic_promotion/unsigned_char_plus_unsigned_char
// origin: languages/c/tests/c/test_lang_arithmetic_promotion.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"250\n"};
int __n = 1, __i = 0;
unsigned char a=200, b=50; { char __t[512]; snprintf(__t, sizeof(__t), "%u\n", (unsigned)(a+b));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

