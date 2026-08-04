// vybe-test: c/stdlib_conversion_and_exit/strtol_empty_returns_zero
// origin: languages/c/tests/c/test_stdlib_conversion_and_exit.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <stdlib.h>
int main() {
const char *__w[] = {"0\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%ld\n", strtol("", 0, 10));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

