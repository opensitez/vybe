// vybe-test: c/output_formatting/printf_table_format
// origin: languages/c/tests/c/test_output_formatting.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"  key      value\n", "    1        100\n"};
int __n = 2, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%5s %10s\n", "key", "value");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
{ char __t[512]; snprintf(__t, sizeof(__t), "%5d %10d\n", 1, 100);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

