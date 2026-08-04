// vybe-test: c/c_string_conversion_strtol_strtod/strtol_with_endptr
// origin: languages/c/tests/c/test_c_string_conversion_strtol_strtod.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int main() {const char *__w[] = {"abc"};
int __n = 1, __i = 0;
 char *end; strtol("123abc", &end, 10); { char __t[512]; snprintf(__t, sizeof(__t), "%s", end);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

