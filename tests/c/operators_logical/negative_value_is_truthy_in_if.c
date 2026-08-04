// vybe-test: c/operators_logical/negative_value_is_truthy_in_if
// origin: languages/c/tests/c/test_operators_logical.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"true\n"};
int __n = 1, __i = 0;
if (-1) { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "true");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } else { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", "false");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

