// vybe-test: c/lang_arithmetic_promotion/short_plus_int_promotes_short
// origin: languages/c/tests/c/test_lang_arithmetic_promotion.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"150\n"};
int __n = 1, __i = 0;
short s=100; int n=50; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", s+n);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

