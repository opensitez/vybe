// vybe-test: c/lang_arithmetic_promotion/int_plus_long_promotes_to_long
// origin: languages/c/tests/c/test_lang_arithmetic_promotion.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"3000\n"};
int __n = 1, __i = 0;
int n=1000; long L=2000; { char __t[512]; snprintf(__t, sizeof(__t), "%ld\n", n+L);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

