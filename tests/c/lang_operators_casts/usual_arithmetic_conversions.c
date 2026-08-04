// vybe-test: c/lang_operators_casts/usual_arithmetic_conversions
// origin: languages/c/tests/c/test_lang_operators_casts.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
int main() {
const char *__w[] = {"4.5\n"};
int __n = 1, __i = 0;
int i=2; double d=2.5; { char __t[512]; snprintf(__t, sizeof(__t), "%.1f\n", i+d);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

