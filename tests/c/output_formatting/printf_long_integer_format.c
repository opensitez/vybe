// vybe-test: c/output_formatting/printf_long_integer_format
// origin: languages/c/tests/c/test_output_formatting.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"1234567890\n"};
int __n = 1, __i = 0;
long x = 1234567890L;
{ char __t[512]; snprintf(__t, sizeof(__t), "%ld\n", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

