// vybe-test: c/ternary_comma/ternary_expression_can_be_assigned_to_variable
// origin: languages/c/tests/c/test_ternary_comma.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"2\n"};
int __n = 1, __i = 0;
int x = 0; int y = x ? 1 : 2; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", y);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

