// vybe-test: c/lang_arithmetic_promotion/short_multiply_int
// origin: languages/c/tests/c/test_lang_arithmetic_promotion.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"36\n"};
int __n = 1, __i = 0;
short s=12; int n=3; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", s*n);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

