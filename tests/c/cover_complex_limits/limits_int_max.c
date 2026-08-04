// vybe-test: c/cover_complex_limits/limits_int_max
// origin: languages/c/tests/c/test_cover_complex_limits.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <limits.h>
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", INT_MAX > 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

