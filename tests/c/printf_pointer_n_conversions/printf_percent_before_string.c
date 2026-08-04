// vybe-test: c/printf_pointer_n_conversions/printf_percent_before_string
// origin: languages/c/tests/c/test_printf_pointer_n_conversions.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <stddef.h>
int main() {
const char *__w[] = {"%s is literal\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%%s is literal\n");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

