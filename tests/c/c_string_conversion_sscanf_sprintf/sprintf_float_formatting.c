// vybe-test: c/c_string_conversion_sscanf_sprintf/sprintf_float_formatting
// origin: languages/c/tests/c/test_c_string_conversion_sscanf_sprintf.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"3.14"};
int __n = 1, __i = 0;
 char buf[50]; sprintf(buf, "%.2f", 3.14159); { char __t[512]; snprintf(__t, sizeof(__t), "%s", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

