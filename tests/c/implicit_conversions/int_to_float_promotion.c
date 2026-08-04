// vybe-test: c/implicit_conversions/int_to_float_promotion
// origin: languages/c/tests/c/test_implicit_conversions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"5.0\n"};
int __n = 1, __i = 0;
int i = 5;
float f = i;
{ char __t[512]; snprintf(__t, sizeof(__t), "%.1f\n", f);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

