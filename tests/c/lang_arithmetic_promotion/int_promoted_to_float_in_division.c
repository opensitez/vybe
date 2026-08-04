// vybe-test: c/lang_arithmetic_promotion/int_promoted_to_float_in_division
// origin: languages/c/tests/c/test_lang_arithmetic_promotion.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"3.5\n"};
int __n = 1, __i = 0;
int n=7; float f=2.0f; { char __t[512]; snprintf(__t, sizeof(__t), "%.1f\n", n/f);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

