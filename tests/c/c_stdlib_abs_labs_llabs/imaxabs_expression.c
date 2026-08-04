// vybe-test: c/c_stdlib_abs_labs_llabs/imaxabs_expression
// origin: languages/c/tests/c/test_c_stdlib_abs_labs_llabs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <inttypes.h>
int main() {const char *__w[] = {"10"};
int __n = 1, __i = 0;
 intmax_t x = 10, y = 20; { char __t[512]; snprintf(__t, sizeof(__t), "%jd", imaxabs(x - y));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

